Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
' StreamJsonRpcMCPClient 和 MCPToolInfo 在根命名空间

Namespace Agent

    ''' <summary>
    ''' 工具描述符 - 描述一个可调用工具的结构
    ''' </summary>
    Public Class ToolDescriptor
        Public Property Id As String
        Public Property Name As String
        Public Property Description As String
        Public Property AppType As String              ' "excel" / "word" / "powerpoint" / "common"
        Public Property Category As String             ' "基础操作" / "数据操作" / "高级功能"
        Public Property RiskLevel As String = "safe"   ' "safe" / "medium" / "risky"
        Public Property IsVbaFallback As Boolean = False
        Public Property Parameters As New List(Of ToolParam)()
    End Class

    ''' <summary>
    ''' 工具参数描述
    ''' </summary>
    Public Class ToolParam
        Public Property Name As String
        Public Property Type As String              ' "string" / "integer" / "boolean" / "array" / "object"
        Public Property Required As Boolean = False
        Public Property Description As String
        Public Property DefaultValue As Object = Nothing
    End Class

    ''' <summary>
    ''' 工具调用结果
    ''' </summary>
    Public Class ToolResult
        Public Property Success As Boolean
        Public Property Message As String
        Public Property Data As Object              ' 执行结果的原始数据
        Public Property ToolId As String
        Public Property ElapsedMs As Long

        Public Shared Function Succeed(toolId As String, Optional message As String = "",
                                       Optional data As Object = Nothing) As ToolResult
            Return New ToolResult With {
                .Success = True,
                .ToolId = toolId,
                .Message = message,
                .Data = data
            }
        End Function

        Public Shared Function Failed(toolId As String, message As String,
                                      Optional data As Object = Nothing) As ToolResult
            Return New ToolResult With {
                .Success = False,
                .ToolId = toolId,
                .Message = message,
                .Data = data
            }
        End Function
    End Class

    ''' <summary>
    ''' 工具调用请求
    ''' </summary>
    Public Class ToolCall
        Public Property ToolId As String
        Public Property Parameters As JObject
        Public Property RequiresApproval As Boolean = False
    End Class

    ''' <summary>
    ''' 工具注册表 - 统一管理 MCP 工具和原生 Office 命令
    ''' </summary>
    Public Class ToolRegistry
        Private ReadOnly _tools As New Dictionary(Of String, ToolDescriptor)(StringComparer.OrdinalIgnoreCase)
        Private _mcpClient As StreamJsonRpcMCPClient
        Private ReadOnly _executeCodeCallback As Action(Of String, String, Boolean)

        ''' <summary>
        ''' 代码执行委托（用于原生 Office 工具）
        ''' </summary>
        Public Property ExecuteCode As Action(Of String, String, Boolean)

        ''' <summary>
        ''' MCP 客户端（用于远程工具调用）
        ''' </summary>
        Public Property McpClient As StreamJsonRpcMCPClient
            Get
                Return _mcpClient
            End Get
            Set(value As StreamJsonRpcMCPClient)
                _mcpClient = value
            End Set
        End Property

        Public Sub New(Optional mcpClient As StreamJsonRpcMCPClient = Nothing)
            _mcpClient = mcpClient
        End Sub

        ''' <summary>
        ''' 从目录加载原生工具定义（JSON 文件）
        ''' </summary>
        Public Sub LoadFromDirectory(dir As String)
            If Not Directory.Exists(dir) Then Return
            For Each file In Directory.GetFiles(dir, "*.json", SearchOption.AllDirectories)
                Try
                    Dim json = System.IO.File.ReadAllText(file)
                    Dim tool = JsonConvert.DeserializeObject(Of ToolDescriptor)(json)
                    If tool IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(tool.Id) Then
                        _tools(tool.Id) = tool
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"[ToolRegistry] 加载工具失败 {file}: {ex.Message}")
                End Try
            Next
        End Sub

        ''' <summary>
        ''' 注册单个工具
        ''' </summary>
        Public Sub RegisterTool(tool As ToolDescriptor)
            _tools(tool.Id) = tool
        End Sub

        ''' <summary>
        ''' 获取指定应用可用的工具
        ''' </summary>
        Public Function GetAvailableTools(appType As String) As List(Of ToolDescriptor)
            Dim result = _tools.Values.Where(Function(t)
                Return String.Equals(t.AppType, appType, StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(t.AppType, "common", StringComparison.OrdinalIgnoreCase)
            End Function).ToList()
            Return result
        End Function

        ''' <summary>
        ''' 获取工具描述
        ''' </summary>
        Public Function GetTool(toolId As String) As ToolDescriptor
            If _tools.ContainsKey(toolId) Then Return _tools(toolId)
            Return Nothing
        End Function

        ''' <summary>
        ''' 检查工具是否存在
        ''' </summary>
        Public Function HasTool(toolId As String) As Boolean
            Return _tools.ContainsKey(toolId)
        End Function

        ''' <summary>
        ''' 自动生成工具描述文本（注入 LLM Prompt）
        ''' </summary>
        Public Function GenerateToolDescriptions(appType As String) As String
            Dim tools = GetAvailableTools(appType)
            Dim sb As New StringBuilder()
            sb.AppendLine($"【已注册工具 - 共 {tools.Count} 个】")
            sb.AppendLine()

            Dim grouped = tools.GroupBy(Function(t) t.Category).OrderBy(Function(g) g.Key)
            For Each group In grouped
                sb.AppendLine($"=== {group.Key} ({group.Count()}个) ===")
                For Each tool In group.OrderBy(Function(t) t.Id)
                    sb.AppendLine($"{tool.Id} - {tool.Name}: {tool.Description}")
                    For Each param In tool.Parameters
                        Dim reqMark = If(param.Required, "必需", "可选")
                        Dim defaultHint = If(param.DefaultValue IsNot Nothing, $", 默认: {param.DefaultValue}", "")
                        sb.AppendLine($"  - {param.Name}({param.Type}, {reqMark}{defaultHint}): {param.Description}")
                    Next
                    sb.AppendLine()
                Next
            Next

            sb.AppendLine()
            sb.AppendLine("【命令格式要求】")
            sb.AppendLine("每个步骤的 code 字段必须是完整 JSON 对象字符串，格式如下：")
            sb.AppendLine("单命令: {""command"":""命令名"",""params"":{...}}")
            sb.AppendLine("多命令: {""commands"":[{""command"":""命令名"",""params"":{...}},...]}")
            sb.AppendLine()
            sb.AppendLine("【绝对禁止】")
            sb.AppendLine("- 禁止使用 actions/operations 数组")
            sb.AppendLine("- 禁止省略 params 包装")
            sb.AppendLine("- 禁止自创未注册的命令")
            sb.AppendLine("- 禁止返回不带代码块的裸 JSON")

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 从 MCP 服务器加载远程工具
        ''' </summary>
        Public Async Function LoadMcpToolsAsync() As Task
            If _mcpClient Is Nothing OrElse Not _mcpClient.IsInitialized Then
                Debug.WriteLine("[ToolRegistry] MCP 客户端未初始化，跳过加载远程工具")
                Return
            End If

            Try
                Dim mcpTools = Await _mcpClient.ListToolsAsync()
                If mcpTools Is Nothing Then Return

                For Each mcpTool In mcpTools
                    If String.IsNullOrWhiteSpace(mcpTool.Name) Then Continue For
                    Dim descriptor = ConvertMcpToDescriptor(mcpTool)
                    _tools(descriptor.Id) = descriptor
                Next

                Debug.WriteLine($"[ToolRegistry] 从 MCP 服务器加载了 {mcpTools.Count} 个工具")
            Catch ex As Exception
                Debug.WriteLine($"[ToolRegistry] 加载 MCP 工具失败: {ex.Message}")
            End Try
        End Function

        ''' <summary>
        ''' 将 MCP 工具信息转换为 ToolDescriptor
        ''' </summary>
        Private Function ConvertMcpToDescriptor(mcpTool As MCPToolInfo) As ToolDescriptor
            Dim descriptor As New ToolDescriptor With {
                .Id = $"mcp.{mcpTool.Name}",
                .Name = mcpTool.Name,
                .Description = If(mcpTool.Description, $"MCP 工具: {mcpTool.Name}"),
                .AppType = "common",
                .Category = "MCP 工具",
                .RiskLevel = "medium"
            }

            ' 解析 InputSchema 中的参数
            Try
                If mcpTool.InputSchema IsNot Nothing Then
                    Dim schema = JObject.FromObject(mcpTool.InputSchema)
                    Dim props = TryCast(schema("properties"), JObject)
                    Dim requiredArray = TryCast(schema("required"), JArray)
                    Dim requiredSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    If requiredArray IsNot Nothing Then
                        For Each r In requiredArray
                            requiredSet.Add(r.ToString())
                        Next
                    End If

                    If props IsNot Nothing Then
                        For Each prop In props.Properties()
                            Dim paramType = "string"
                            Dim propType = prop.Value("type")?.ToString()
                            If Not String.IsNullOrEmpty(propType) Then
                                Select Case propType.ToLower()
                                    Case "integer", "number"
                                        paramType = "integer"
                                    Case "boolean"
                                        paramType = "boolean"
                                    Case "array"
                                        paramType = "array"
                                    Case "object"
                                        paramType = "object"
                                    Case Else
                                        paramType = "string"
                                End Select
                            End If

                            descriptor.Parameters.Add(New ToolParam With {
                                .Name = prop.Name,
                                .Type = paramType,
                                .Required = requiredSet.Contains(prop.Name),
                                .Description = prop.Value("description")?.ToString()
                            })
                        Next
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine($"[ToolRegistry] 解析 MCP 工具参数失败 {mcpTool.Name}: {ex.Message}")
            End Try

            Return descriptor
        End Function

        ''' <summary>
        ''' 执行工具调用
        ''' </summary>
        Public Async Function ExecuteToolAsync(toolId As String, params As JObject) As Task(Of ToolResult)
            Dim sw = Diagnostics.Stopwatch.StartNew()

            Dim tool = GetTool(toolId)
            If tool Is Nothing Then
                sw.Stop()
                Return ToolResult.Failed(toolId, $"未找到工具: {toolId}")
            End If

            ' MCP 工具调用（以 mcp. 开头）
            If toolId.StartsWith("mcp.") Then
                If _mcpClient Is Nothing OrElse Not _mcpClient.IsInitialized Then
                    sw.Stop()
                    Return ToolResult.Failed(toolId, "MCP 客户端未初始化")
                End If

                Try
                    Dim actualToolName = toolId.Substring(4)
                    Dim mcpResult = Await _mcpClient.CallToolAsync(actualToolName, params)
                    sw.Stop()

                    If mcpResult.IsError Then
                        Return ToolResult.Failed(toolId, If(mcpResult.ErrorMessage, "MCP 工具执行失败"))
                    End If

                    Dim outputText As String = ""
                    If mcpResult.Content IsNot Nothing AndAlso mcpResult.Content.Count > 0 Then
                        Dim sb As New StringBuilder()
                        For Each content In mcpResult.Content
                            If content.Type = "text" AndAlso Not String.IsNullOrEmpty(content.Text) Then
                                sb.AppendLine(content.Text)
                            End If
                        Next
                        outputText = sb.ToString().Trim()
                    End If

                    Return ToolResult.Succeed(toolId, If(String.IsNullOrEmpty(outputText), "执行成功", outputText),
                                               New With {.elapsedMs = sw.ElapsedMilliseconds})
                Catch ex As Exception
                    sw.Stop()
                    Return ToolResult.Failed(toolId, $"MCP 调用异常: {ex.Message}")
                End Try
            End If

            ' 原生 Office 工具，通过 ExecuteCode 回调执行
            If tool.IsVbaFallback OrElse Not toolId.StartsWith("mcp.") Then
                If ExecuteCode Is Nothing Then
                    sw.Stop()
                    Return ToolResult.Failed(toolId, "ExecuteCode 回调未设置")
                End If

                ' 构建完整的 JSON 命令
                Dim command As String
                If params.ContainsKey("commands") Then
                    command = params.ToString(Formatting.None)
                Else
                    Dim wrapped = New JObject From {
                        {"command", toolId},
                        {"params", params}
                    }
                    command = wrapped.ToString(Formatting.None)
                End If

                Try
                    ' 调用现有执行逻辑
                    ExecuteCode.Invoke(command, "json", False)
                    sw.Stop()
                    Return ToolResult.Succeed(toolId, "执行成功", New With {.elapsedMs = sw.ElapsedMilliseconds})
                Catch ex As Exception
                    sw.Stop()
                    Return ToolResult.Failed(toolId, $"执行失败: {ex.Message}")
                End Try
            End If

            sw.Stop()
            Return ToolResult.Failed(toolId, "未知的工具类型")
        End Function

        ''' <summary>
        ''' 获取所有已注册工具数量
        ''' </summary>
        Public ReadOnly Property ToolCount As Integer
            Get
                Return _tools.Count
            End Get
        End Property

        ''' <summary>
        ''' 清空所有工具
        ''' </summary>
        Public Sub Clear()
            _tools.Clear()
        End Sub
    End Class

End Namespace
