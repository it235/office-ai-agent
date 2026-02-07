Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Converters
Imports Newtonsoft.Json.Linq
Imports ShareRibbon.MCPConnectionManager

' MCP连接配置类，完全对应官方格式
Public Class MCPConnectionConfig
    ' HTTP/SSE 连接字段
    <JsonProperty(PropertyName:="name")>
    Public Property Name As String

    <JsonProperty(PropertyName:="description")>
    Public Property Description As String

    <JsonProperty(PropertyName:="isActive")>
    Public Property IsActive As Boolean

    <JsonProperty(PropertyName:="baseUrl")>
    Public Property BaseUrl As String

    ' Stdio 连接字段
    <JsonProperty(PropertyName:="command")>
    Public Property Command As String

    <JsonProperty(PropertyName:="args")>
    Public Property Args As List(Of String)

    <JsonProperty(PropertyName:="env")>
    Public Property Env As Dictionary(Of String, String)

    ' 额外保留的字段 (不存储到官方格式)
    <JsonProperty(PropertyName:="tools", NullValueHandling:=NullValueHandling.Ignore)>
    Public Property Tools As List(Of JObject)


    ' 兼容属性 - 用于兼容旧的代码逻辑
    <JsonIgnore>
    Public Property Enabled As Boolean
        Get
            Return IsActive
        End Get
        Set(value As Boolean)
            IsActive = value
        End Set
    End Property

    <JsonIgnore>
    Public Property ConnectionType As String
        Get
            Return If(IsStdio, "Stdio", "HTTP")
        End Get
        Set(value As String)
            ' 设置连接类型 - 仅用于向后兼容
        End Set
    End Property

    <JsonIgnore>
    Public Property Url As String
        Get
            If IsStdio Then
                ' 构建Stdio URL
                Return GetStdioUrl()
            Else
                Return BaseUrl
            End If
        End Get
        Set(value As String)
            If value.StartsWith("stdio://") Then
                ' 解析Stdio URL
                ParseStdioUrl(value)
            Else
                BaseUrl = value
                Command = Nothing
                Args = New List(Of String)()
                Env = New Dictionary(Of String, String)()
            End If
        End Set
    End Property

    <JsonIgnore>
    Public Property EnvironmentVariables As Dictionary(Of String, String)
        Get
            Return Env
        End Get
        Set(value As Dictionary(Of String, String))
            Env = value
        End Set
    End Property

    <JsonIgnore>
    Public ReadOnly Property IsStdio As Boolean
        Get
            Return Not String.IsNullOrEmpty(Command)
        End Get
    End Property

    Public Sub New()
        Name = String.Empty
        Description = String.Empty
        IsActive = True
        BaseUrl = String.Empty
        Command = Nothing
        Args = New List(Of String)()
        Env = New Dictionary(Of String, String)()
        Tools = New List(Of JObject)()
    End Sub

    Public Sub New(name As String, url As String, connectionType As String)
        Me.Name = name
        Me.Description = String.Empty
        Me.IsActive = True
        Me.Tools = New List(Of JObject)()
        Me.Env = New Dictionary(Of String, String)()
        Me.Args = New List(Of String)()

        ' 设置URL和连接类型
        If connectionType.Equals("Stdio", StringComparison.OrdinalIgnoreCase) Then
            ParseStdioUrl(url)
        Else
            BaseUrl = url
            Command = Nothing
        End If
    End Sub


    ' 解析Stdio URL并设置相关属性
    Private Sub ParseStdioUrl(stdioUrl As String)
        If Not stdioUrl.StartsWith("stdio://") Then
            Return
        End If

        ' 解析URL
        Dim options = StdioOptions.Parse(stdioUrl)
        Command = options.Command

        ' 处理参数字符串为数组
        If Not String.IsNullOrEmpty(options.Arguments) Then
            Args = ParseArgumentsString(options.Arguments)
        Else
            Args = New List(Of String)()
        End If

        ' 复制环境变量
        Env.Clear()
        For Each kvp In options.EnvironmentVariables
            Env.Add(kvp.Key, kvp.Value)
        Next
    End Sub

    ' 生成Stdio URL
    Private Function GetStdioUrl() As String
        If String.IsNullOrEmpty(Command) Then
            Return String.Empty
        End If

        Dim options = New StdioOptions()
        options.Command = Command

        ' 将参数数组转为字符串 - 确保保留原始格式
        If Args IsNot Nothing AndAlso Args.Count > 0 Then
            ' 直接使用Args中的原始路径
            options.Arguments = String.Join(" ", Args.Select(Function(arg) If(arg.Contains(" "), $"""{arg}""", arg)))
        End If

        ' 复制环境变量
        options.EnvironmentVariables = New Dictionary(Of String, String)(Env)

        Return options.ToUrl()
    End Function

    ' 解析命令行参数字符串为数组，处理引号内的空格
    Private Shared Function ParseArgumentsString(argsString As String) As List(Of String)
        Dim result As New List(Of String)()
        Dim currentArg As New StringBuilder()
        Dim inQuotes As Boolean = False
        Dim escaping As Boolean = False

        For i As Integer = 0 To argsString.Length - 1
            Dim c As Char = argsString(i)

            ' 处理转义字符
            If escaping Then
                ' 保留转义字符和被转义的字符
                currentArg.Append(c)
                escaping = False
                Continue For
            End If

            ' 检查是否是转义字符
            If c = "\"c Then
                ' 保留反斜杠，不把它当作转义字符
                currentArg.Append(c)
                Continue For
            End If

            ' 处理引号
            If c = """"c Then
                inQuotes = Not inQuotes
                ' 保留引号，因为某些命令行工具需要引号来处理路径中的空格
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

        Return result
    End Function

    ' 从环境变量表格更新Env属性
    Public Sub UpdateEnvFromGrid(grid As DataGridView)
        Env.Clear()

        For Each row As DataGridViewRow In grid.Rows
            If row.IsNewRow Then Continue For

            Dim key = TryCast(row.Cells("Key").Value, String)
            Dim value = TryCast(row.Cells("Value").Value, String)

            If Not String.IsNullOrEmpty(key) Then
                Env(key) = If(value IsNot Nothing, value, "")
            End If
        Next
    End Sub

    ' 填充环境变量表格
    Public Sub FillEnvGrid(grid As DataGridView)
        grid.Rows.Clear()

        For Each kvp In Env
            Dim rowIndex = grid.Rows.Add()
            grid.Rows(rowIndex).Cells("Key").Value = kvp.Key
            grid.Rows(rowIndex).Cells("Value").Value = kvp.Value
        Next
    End Sub
End Class

' 自定义JSON转换器，确保路径中的反斜杠不被过度转义
Public Class PathPreservingStringConverter
    Inherits JsonConverter

    Public Overrides Function CanConvert(objectType As Type) As Boolean
        Return objectType = GetType(String)
    End Function

    Public Overrides Function ReadJson(reader As JsonReader, objectType As Type, existingValue As Object, serializer As JsonSerializer) As Object
        Return reader.Value
    End Function

    Public Overrides Sub WriteJson(writer As JsonWriter, value As Object, serializer As JsonSerializer)
        Dim str = TryCast(value, String)
        If str IsNot Nothing Then
            ' 直接写入字符串，不做任何转义
            writer.WriteValue(str)
        Else
            writer.WriteNull()
        End If
    End Sub
End Class

Public Class MCPConnectionManager
    Private Shared ReadOnly _connectionsFile As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        ConfigSettings.OfficeAiAppDataFolder,
        "mcp_connection.json")

    ' 官方格式连接配置根类
    Public Class MCPConnectionsRoot
        <JsonProperty(PropertyName:="mcpServers")>
        Public Property McpServers As Dictionary(Of String, MCPConnectionConfig)

        Public Sub New()
            McpServers = New Dictionary(Of String, MCPConnectionConfig)()
        End Sub
    End Class

    ' 加载连接
    Public Shared Function LoadConnections() As List(Of MCPConnectionConfig)
        Try
            ' 确保目录存在
            Dim directoryx = Path.GetDirectoryName(_connectionsFile)
            If Not Directory.Exists(directoryx) Then
                Directory.CreateDirectory(directoryx)
            End If

            ' 如果文件存在，加载配置
            If File.Exists(_connectionsFile) Then
                Dim json = File.ReadAllText(_connectionsFile)

                ' 配置JsonSerializerSettings来处理特殊路径字符
                Dim settings = New JsonSerializerSettings()

                Dim newFormat = JsonConvert.DeserializeObject(Of MCPConnectionsRoot)(json, settings)

                If newFormat?.McpServers IsNot Nothing Then
                    ' 转换为列表返回
                    Dim connectionsList = New List(Of MCPConnectionConfig)()

                    For Each kvp In newFormat.McpServers
                        Dim config = kvp.Value
                        ' 记录服务器ID作为名称(如果名称为空)
                        If String.IsNullOrEmpty(config.Name) Then
                            config.Name = kvp.Key
                        End If

                        connectionsList.Add(config)
                    Next

                    Return connectionsList
                End If
            End If

            ' 如果文件不存在或解析失败，返回空列表
            Return New List(Of MCPConnectionConfig)()
        Catch ex As Exception
            Debug.WriteLine($"加载连接配置失败: {ex.Message}")
            Return New List(Of MCPConnectionConfig)()
        End Try
    End Function

    ' 导入并合并配置方法
    ' 导入并合并配置方法
    Public Shared Function ImportAndMergeConfig(jsonConfig As String) As Integer
        Try
            ' 解析导入的JSON
            Dim importedConfig = JObject.Parse(jsonConfig)

            ' 检查是否包含mcpServers节点
            If importedConfig("mcpServers") Is Nothing Then
                Return 0
            End If

            ' 获取mcpServers对象
            Dim importedServers = importedConfig("mcpServers").ToObject(Of JObject)()

            ' 读取现有配置
            Dim existingConfig As MCPConnectionsRoot = Nothing

            ' 确保目录存在
            Dim directoryx = Path.GetDirectoryName(_connectionsFile)
            If Not Directory.Exists(directoryx) Then
                Directory.CreateDirectory(directoryx)
            End If

            ' 检查现有配置文件
            If File.Exists(_connectionsFile) Then
                Try
                    Dim existingJson = File.ReadAllText(_connectionsFile)
                    existingConfig = JsonConvert.DeserializeObject(Of MCPConnectionsRoot)(existingJson)
                Catch ex As Exception
                    Debug.WriteLine($"读取现有配置失败，将创建新配置: {ex.Message}")
                    existingConfig = New MCPConnectionsRoot()
                End Try
            Else
                existingConfig = New MCPConnectionsRoot()
            End If

            ' 如果McpServers为null，初始化它
            If existingConfig.McpServers Is Nothing Then
                existingConfig.McpServers = New Dictionary(Of String, MCPConnectionConfig)()
            End If

            ' 导入的服务器数量
            Dim importedCount = 0

            ' 遍历导入的服务器配置并合并
            For Each serverProp In importedServers.Properties()
                Dim serverId = serverProp.Name
                Dim serverConfig = serverProp.Value

                ' 转换为MCPConnectionConfig对象
                Dim connectionConfig As MCPConnectionConfig = Nothing

                Try
                    ' 尝试直接反序列化
                    connectionConfig = serverConfig.ToObject(Of MCPConnectionConfig)()

                    ' 确保名称不为空
                    If String.IsNullOrEmpty(connectionConfig.Name) Then
                        connectionConfig.Name = serverId
                    End If

                    ' 确保描述字段存在
                    If connectionConfig.Description Is Nothing Then
                        connectionConfig.Description = String.Empty
                    End If

                    ' 设置激活状态
                    If connectionConfig.IsActive = Nothing Then
                        connectionConfig.IsActive = True
                    End If

                    ' 初始化工具列表
                    connectionConfig.Tools = New List(Of JObject)()

                    ' 添加到现有配置
                    existingConfig.McpServers(serverId) = connectionConfig
                    importedCount += 1
                Catch ex As Exception
                    Debug.WriteLine($"无法导入服务器 {serverId}: {ex.Message}")
                    ' 继续处理下一个
                    Continue For
                End Try
            Next

            ' 如果成功导入了配置，保存更新后的配置
            If importedCount > 0 Then
                ' 序列化并保存
                Dim settings = New JsonSerializerSettings() With {
                .Formatting = Formatting.Indented
            }
                Dim json = JsonConvert.SerializeObject(existingConfig, settings)
                File.WriteAllText(_connectionsFile, json)
            End If

            Return importedCount
        Catch ex As Exception
            Debug.WriteLine($"导入配置失败: {ex.Message}")
            Throw
        End Try
    End Function
    ' 保存连接
    Public Shared Function SaveConnections(connections As List(Of MCPConnectionConfig)) As Boolean
        Try
            ' 确保目录存在
            Dim directoryx = Path.GetDirectoryName(_connectionsFile)
            If Not Directory.Exists(directoryx) Then
                Directory.CreateDirectory(directoryx)
            End If

            ' 将列表转换为官方格式
            Dim newFormat = New MCPConnectionsRoot()

            For Each config In connections
                ' 生成唯一标识符
                Dim serverId As String
                If config.IsStdio Then
                    ' Stdio连接使用命令作为ID基础
                    serverId = GenerateStdioId(config.Command)
                Else
                    ' HTTP连接使用唯一ID
                    serverId = GenerateUniqueId(config.Name)
                End If

                ' 添加到字典
                newFormat.McpServers(serverId) = config
            Next

            ' 使用特殊设置序列化，确保路径正确保存
            Dim settings = New JsonSerializerSettings() With {
                .Formatting = Formatting.Indented
            }

            ' 创建自定义序列化器，让JSON.NET使用原始字符串，不做任何转义
            settings.Converters.Add(New StringEnumConverter())

            ' 这种方法太复杂了，我们直接使用JObject来保存，手动处理反斜杠
            Dim x = JObject.FromObject(newFormat)

            ' 修复字符串中的反斜杠 - 直接用字符串替换处理
            Dim json = x.ToString(Newtonsoft.Json.Formatting.Indented)

            ' 处理Command路径中的反斜杠
            For Each server In newFormat.McpServers
                If server.Value.IsStdio Then
                    ' 替换原有JSON中的路径，确保使用双反斜杠表示
                    json = json.Replace($"""{server.Value.Command}""", $"""{server.Value.Command.Replace("\", "\\")}""")

                    ' 处理参数中的反斜杠
                    If server.Value.Args IsNot Nothing Then
                        For Each arg In server.Value.Args
                            If arg.Contains("\") Then
                                ' 替换参数中的反斜杠
                                json = json.Replace($"""{arg}""", $"""{arg.Replace("\", "\\")}""")
                            End If
                        Next
                    End If
                End If
            Next

            ' 保存处理后的JSON
            File.WriteAllText(_connectionsFile, json)

            Return True
        Catch ex As Exception
            Debug.WriteLine($"保存连接配置失败: {ex.Message}")
            Return False
        End Try
    End Function

    ' 根据Stdio命令生成ID
    Private Shared Function GenerateStdioId(command As String) As String
        ' 从命令中提取基本名称，移除路径
        Dim baseName = Path.GetFileNameWithoutExtension(command)

        ' 清理ID，只保留字母数字和连字符
        Dim cleanId = New String(baseName.Where(Function(c) Char.IsLetterOrDigit(c) OrElse c = "-"c).ToArray())

        ' 如果为空，使用默认值
        If String.IsNullOrEmpty(cleanId) Then
            cleanId = "stdio-server"
        End If

        ' 确保唯一性
        Dim uniqueId = cleanId & "-" & DateTime.Now.ToString("yyyyMMddHHmmss")
        Return uniqueId.ToLowerInvariant()
    End Function

    ' 根据名称生成唯一ID
    Private Shared Function GenerateUniqueId(name As String) As String
        ' 生成唯一ID
        Return Guid.NewGuid().ToString("N").Substring(0, 12)
    End Function

    ' 移除连接
    Public Shared Function RemoveConnection(connections As List(Of MCPConnectionConfig), connectionName As String) As List(Of MCPConnectionConfig)
        ' 查找并移除连接
        connections.RemoveAll(Function(c) c.Name.Equals(connectionName, StringComparison.OrdinalIgnoreCase))

        ' 保存并返回更新后的列表
        SaveConnections(connections)
        Return connections
    End Function
End Class