' ShareRibbon\Controls\Services\McpService.vb
' MCP 服务：处理 MCP 连接管理和工具调用

Imports System.Linq
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' MCP 服务，负责 MCP 连接管理和设置
''' </summary>
Public Class McpService

    Private ReadOnly _executeScript As Func(Of String, Threading.Tasks.Task)
    Private ReadOnly _getApplication As Func(Of ApplicationInfo)

    ''' <summary>
    ''' 构造函数
    ''' </summary>
    Public Sub New(executeScript As Func(Of String, Threading.Tasks.Task), getApplication As Func(Of ApplicationInfo))
        _executeScript = executeScript
        _getApplication = getApplication
    End Sub

    ''' <summary>
    ''' 获取 MCP 连接列表并发送到前端
    ''' </summary>
    Public Sub GetMcpConnections()
        Try
            Dim connections = MCPConnectionManager.LoadConnections()
            Dim enabledConnections = connections.Where(Function(c) c.IsActive).ToList()

            Dim chatSettings As New ChatSettings(_getApplication())
            Dim enabledMcpList = chatSettings.EnabledMcpList

            Dim connectionsJson = JsonConvert.SerializeObject(enabledConnections)
            Dim enabledListJson = JsonConvert.SerializeObject(enabledMcpList)
            Dim mcpSupported As Boolean = ConfigSettings.mcpable

            Dim js = $"renderMcpConnections({connectionsJson}, {enabledListJson},{mcpSupported.ToString().ToLower()});"
            _executeScript(js)
        Catch ex As Exception
            Debug.WriteLine($"获取MCP连接列表失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 保存 MCP 设置
    ''' </summary>
    Public Sub SaveMcpSettings(jsonDoc As JObject)
        Try
            Dim enabledList As List(Of String) = jsonDoc("enabledList").ToObject(Of List(Of String))()
            Dim chatSettings As New ChatSettings(_getApplication())
            chatSettings.SaveEnabledMcpList(enabledList)
            GlobalStatusStrip.ShowInfo("MCP设置已保存")
        Catch ex As Exception
            Debug.WriteLine($"保存MCP设置失败: {ex.Message}")
            GlobalStatusStrip.ShowWarning("保存MCP设置失败")
        End Try
    End Sub

    ''' <summary>
    ''' 初始化 MCP 设置
    ''' </summary>
    Public Sub InitializeMcpSettings()
        Try
            Dim mcpSupported = False
            For Each config In ConfigManager.ConfigData
                If config.selected Then
                    For Each model In config.model
                        If model.selected Then
                            mcpSupported = model.mcpable
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next

            ' 加载MCP连接和启用列表
            Dim connections = MCPConnectionManager.LoadConnections()
            Dim enabledConnections = connections.Where(Function(c) c.IsActive).ToList()

            Dim chatSettings As New ChatSettings(_getApplication())
            Dim enabledMcpList = chatSettings.EnabledMcpList

            Dim connectionsJson = JsonConvert.SerializeObject(enabledConnections)
            Dim enabledListJson = JsonConvert.SerializeObject(enabledMcpList)

            Dim js = $"setMcpSupport({mcpSupported.ToString().ToLower()}, {connectionsJson}, {enabledListJson});"
            _executeScript(js)
        Catch ex As Exception
            Debug.WriteLine($"初始化MCP设置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 创建错误响应
    ''' </summary>
    Public Shared Function CreateErrorResponse(errorMessage As String) As JObject
        Dim responseObj = New JObject()
        responseObj("isError") = True
        responseObj("errorMessage") = errorMessage
        responseObj("timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Debug.WriteLine($"创建错误响应: {errorMessage}")
        Return responseObj
    End Function

End Class
