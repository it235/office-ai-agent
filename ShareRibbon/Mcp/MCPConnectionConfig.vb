Imports System.Collections.Generic
Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class MCPConnectionConfig
    Public Property Name As String
    Public Property Description As String
    Public Property Url As String
    Public Property ConnectionType As String  ' "HTTP" 或 "Stdio"
    Public Property Enabled As Boolean
    Public Property LastConnected As DateTime
    Public Property EnvironmentVariables As Dictionary(Of String, String)
    Public Property Tools As List(Of JObject)  ' 新增: 存储工具列表，兼容大模型函数调用格式

    Public Sub New()
        Enabled = True
        LastConnected = DateTime.MinValue
        Description = String.Empty
        EnvironmentVariables = New Dictionary(Of String, String)()
        Tools = New List(Of JObject)()  ' 初始化工具列表
    End Sub

    Public Sub New(name As String, url As String, connectionType As String)
        Me.Name = name
        Me.Description = String.Empty  ' 初始化描述字段
        Me.Url = url
        Me.ConnectionType = connectionType
        Me.Enabled = True
        Me.LastConnected = DateTime.MinValue
        Me.EnvironmentVariables = New Dictionary(Of String, String)()
        Me.Tools = New List(Of JObject)()  ' 初始化工具列表
    End Sub
End Class

Public Class MCPConnectionManager
    Private Shared ReadOnly _connectionsFile As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        ConfigSettings.OfficeAiAppDataFolder,
        "mcp_connections.json")

    Public Shared Function LoadConnections() As List(Of MCPConnectionConfig)
        Try
            ' 确保目录存在
            Dim directoryx = Path.GetDirectoryName(_connectionsFile)
            If Not Directory.Exists(directoryx) Then
                Directory.CreateDirectory(directoryx)
            End If

            ' 如果文件存在，加载连接
            If File.Exists(_connectionsFile) Then
                Dim json = File.ReadAllText(_connectionsFile)
                Return JsonConvert.DeserializeObject(Of List(Of MCPConnectionConfig))(json)
            End If

            ' 如果文件不存在，返回空列表
            Return New List(Of MCPConnectionConfig)()
        Catch ex As Exception
            Debug.WriteLine($"加载连接配置失败: {ex.Message}")
            Return New List(Of MCPConnectionConfig)()
        End Try
    End Function

    Public Shared Function SaveConnections(connections As List(Of MCPConnectionConfig)) As Boolean
        Try
            ' 确保目录存在
            Dim directoryx = Path.GetDirectoryName(_connectionsFile)
            If Not directory.Exists(directoryx) Then
                directory.CreateDirectory(directoryx)
            End If

            ' 序列化并保存
            Dim json = JsonConvert.SerializeObject(connections, Formatting.Indented)
            File.WriteAllText(_connectionsFile, json)
            Return True
        Catch ex As Exception
            Debug.WriteLine($"保存连接配置失败: {ex.Message}")
            Return False
        End Try
    End Function

    Public Shared Function RemoveConnection(connections As List(Of MCPConnectionConfig), connectionName As String) As List(Of MCPConnectionConfig)
        ' 查找并移除连接
        connections.RemoveAll(Function(c) c.Name.Equals(connectionName, StringComparison.OrdinalIgnoreCase))

        ' 保存并返回更新后的列表
        SaveConnections(connections)
        Return connections
    End Function
End Class