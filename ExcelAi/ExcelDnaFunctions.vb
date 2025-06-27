Imports System.Diagnostics
Imports System.Threading.Tasks
Imports ExcelDna.Integration
Imports Newtonsoft.Json.Linq
Imports ShareRibbon

Public Module ExcelDnaFunctions

    ' 配置管理器实例
    Private configManager As ConfigManager = Nothing

    ' 修复AutoOpen方法 - 移除MenuName和MenuText参数
    <ExcelDna.Integration.ExcelCommand()>
    Public Sub AutoOpen()
        Try
            Debug.WriteLine("=== Excel-DNA AutoOpen 开始执行 ===")
            Console.WriteLine("=== Excel-DNA AutoOpen 开始执行 ===")

            ' 在Excel-DNA域中初始化GlobalStatusStrip
            Try
                Dim excelApp As Object = ExcelDna.Integration.ExcelDnaUtil.Application
                If excelApp IsNot Nothing Then
                    GlobalStatusStripAll.InitializeApplication(excelApp)
                    Debug.WriteLine("Excel-DNA域中GlobalStatusStrip初始化成功")

                    ' 测试状态栏是否工作
                    Try
                        excelApp.StatusBar = "Excel-DNA已加载"
                        Debug.WriteLine("Excel-DNA状态栏测试成功")
                        Threading.Thread.Sleep(2000)
                        excelApp.StatusBar = False
                    Catch statusEx As Exception
                        Debug.WriteLine($"Excel-DNA状态栏测试失败: {statusEx.Message}")
                    End Try
                Else
                    Debug.WriteLine("Excel-DNA域中无法获取Excel应用程序对象")
                End If
            Catch ex As Exception
                Debug.WriteLine($"Excel-DNA域中GlobalStatusStrip初始化失败: {ex.Message}")
            End Try

            ' 初始化配置管理器
            Try
                configManager = New ConfigManager()
                configManager.LoadConfig()
                Debug.WriteLine($"Excel-DNA 配置已加载 - API URL: {ConfigSettings.ApiUrl}")
                Debug.WriteLine($"Excel-DNA 配置已加载 - Model: {ConfigSettings.ModelName}")
            Catch ex As Exception
                Debug.WriteLine($"Excel-DNA 配置加载失败: {ex.Message}")
            End Try

            ' 清除缓存
            FunctionCache.Clear()

            ' 注册未处理异常处理程序
            ExcelDna.Integration.ExcelIntegration.RegisterUnhandledExceptionHandler(
            Function(ex) $"错误: {ex.Message}")

            Debug.WriteLine("=== Excel-DNA AutoOpen 执行完成 ===")
            Console.WriteLine("=== Excel-DNA AutoOpen 执行完成 ===")

        Catch ex As Exception
            Debug.WriteLine($"Excel-DNA AutoOpen 执行失败: {ex.Message}")
            Console.WriteLine($"Excel-DNA AutoOpen 执行失败: {ex.Message}")
        End Try
    End Sub


    ' 确保配置已加载
    Private Sub EnsureConfigLoaded()
        If configManager Is Nothing Then
            configManager = New ConfigManager()
            configManager.LoadConfig()
        End If
    End Sub

    ' 基本 LLM 函数
    <ExcelFunction(Description:="使用AI模型生成文本",
                  Category:="Excel AI 函数",
                  Name:="ELLM",
                  IsVolatile:=False,
                  IsThreadSafe:=True)>
    Public Function ELLM(
        <ExcelArgument(Description:="提示词或问题")> prompt As String) As Object

        ' 使用缓存避免重复计算
        Dim cacheKey As String = $"ELLM|{prompt}"
        If FunctionCache.ContainsKey(cacheKey) Then
            Return FunctionCache(cacheKey)
        End If

        Try
            ' 调用完整版函数
            Dim result As String = ADLLM(prompt, "", "", 0.7, 1000)

            ' 缓存结果
            If Not result.StartsWith("错误:") Then
                FunctionCache(cacheKey) = result
            End If

            Return result
        Catch ex As Exception
            Return $"错误: {ex.Message}"
        End Try
    End Function

    ' 高级 LLM 函数
    <ExcelFunction(Description:="使用AI模型生成文本(高级版)",
                  Category:="Excel AI 函数",
                  Name:="ADLLM",
                  IsVolatile:=False,
                  IsThreadSafe:=True)>
    Public Function ADLLM(
        <ExcelArgument(Description:="提示词或问题")> prompt As String,
        <ExcelArgument(Description:="可选: 模型名称")> Optional model As String = "",
        <ExcelArgument(Description:="可选: 系统提示词")> Optional systemPrompt As String = "",
        <ExcelArgument(Description:="可选: 温度参数 (0.0-1.0)")> Optional temperature As Double = 0.7,
        <ExcelArgument(Description:="可选: 最大生成令牌数")> Optional maxTokens As Integer = 1000) As Object

        ' 使用缓存避免重复计算
        Dim cacheKey As String = $"ADLLM|{prompt}|{model}|{systemPrompt}|{temperature}|{maxTokens}"
        If FunctionCache.ContainsKey(cacheKey) Then
            Return FunctionCache(cacheKey)
        End If

        Try
            ' 验证输入
            If String.IsNullOrEmpty(prompt) Then
                Return "错误: 提示词不能为空"
            End If

            ' 确保配置已加载
            EnsureConfigLoaded()
            ' 使用默认值（如果未提供）
            Dim apiKey As String = ConfigSettings.ApiKey
            Dim apiUrl As String = ConfigSettings.ApiUrl

            If String.IsNullOrEmpty(apiKey) Then
                Return "错误: 未配置API密钥"
            End If

            If String.IsNullOrEmpty(apiUrl) Then
                Return "错误: 未配置API URL"
            End If

            ' 使用指定的模型或默认模型
            Dim useModel As String = If(String.IsNullOrEmpty(model), GetDefaultModel(), model)

            GlobalStatusStripAll.ShowWarning($"正在同步请求大模型中，请耐心等待")

            ' 创建请求体
            Dim requestBody As String = LLMUtil.CreateLlmRequestBody(prompt, useModel, systemPrompt, temperature, maxTokens)

            ' 调用API并获取结果
            Dim response As String = LLMUtil.SendHttpRequest(apiUrl, apiKey, requestBody).Result

            ' 如果响应为空，返回错误信息
            If String.IsNullOrEmpty(response) Then
                Return "错误: API未返回响应"
            End If

            Dim parsedResponse As JObject = JObject.Parse(response)
            Dim cellValue As String = parsedResponse("choices")(0)("message")("content").ToString()

            ' 缓存结果
            FunctionCache(cacheKey) = cellValue

            Return cellValue
        Catch ex As Exception
            Return $"错误: {ex.Message}"
        End Try
    End Function

    ' 异步 LLM 函数 - 修复版本，去掉未定义的接口
    <ExcelFunction(Description:="异步调用AI模型生成文本",
                  Category:="Excel AI 函数",
                  Name:="ALLM",
                  IsVolatile:=False,
                  IsThreadSafe:=True)>
    Public Function ALLM(
        <ExcelArgument(Description:="提示词或问题")> prompt As String,
        <ExcelArgument(Description:="可选: 模型名称")> Optional model As String = "",
        <ExcelArgument(Description:="可选: 系统提示词")> Optional systemPrompt As String = "",
        <ExcelArgument(Description:="可选: 温度参数 (0.0-1.0)")> Optional temperature As Double = 0.7,
        <ExcelArgument(Description:="可选: 最大生成令牌数")> Optional maxTokens As Integer = 1000) As Object

        If temperature = 0.0 Then temperature = 0.7
        If maxTokens = 0 Then maxTokens = 1000

        ' 使用Excel-DNA特有的线程安全方法设置状态栏
        Try
            Dim displayPrompt As String = If(prompt.Length > 25, prompt.Substring(0, 25) + "...", prompt)
            SetExcelStatusBarDirectly($"正在思考「{displayPrompt}」...")
        Catch ex As Exception
            Debug.WriteLine($"显示状态提示失败: {ex.Message}")
        End Try

        ' 检查缓存，如果有缓存直接返回
        Dim cacheKey As String = $"ALLM|{prompt}|{model}|{systemPrompt}|{temperature}|{maxTokens}"
        If FunctionCache.ContainsKey(cacheKey) Then
            Return FunctionCache(cacheKey)
        End If


        ' 使用 ExcelAsyncUtil.Run 执行异步操作
        Return ExcelAsyncUtil.Run("ALLM", New Object() {prompt, model, systemPrompt, temperature, maxTokens},
            Function()
                Try
                    Dim result As String = ProcessLLMRequestSync(prompt, model, systemPrompt, temperature, maxTokens)

                    ' 清除状态栏
                    ClearExcelStatusBar()
                    Return result
                Catch ex As Exception
                    ClearExcelStatusBar()
                    Return $"错误: {ex.Message}"
                End Try
            End Function)
    End Function

    ' 专门用于Excel-DNA的线程安全状态栏设置
    Private Sub SetExcelStatusBarDirectly(message As String)
        Try
            Debug.WriteLine($"正在尝试安全设置状态栏: {message}")

            ' 使用Excel-DNA的QueueAsMacro确保在Excel主线程上执行
            ExcelAsyncUtil.QueueAsMacro(Sub()
                                            Try
                                                Dim excelApp As Object = ExcelDna.Integration.ExcelDnaUtil.Application
                                                If excelApp IsNot Nothing Then
                                                    excelApp.StatusBar = message
                                                    Debug.WriteLine($"状态栏设置成功: {message}")
                                                Else
                                                    Debug.WriteLine("无法获取Excel应用程序对象")
                                                End If
                                            Catch innerEx As Exception
                                                Debug.WriteLine($"在Excel主线程上设置状态栏失败: {innerEx.Message}")
                                            End Try
                                        End Sub)
        Catch ex As Exception
            Debug.WriteLine($"队列状态栏操作失败: {ex.Message}")
        End Try
    End Sub

    ' 在ProcessLLMRequestSync完成后清除状态栏
    Private Sub ClearExcelStatusBar()
        Try
            ExcelAsyncUtil.QueueAsMacro(Sub()
                                            Try
                                                Dim excelApp As Object = ExcelDna.Integration.ExcelDnaUtil.Application
                                                If excelApp IsNot Nothing Then
                                                    excelApp.StatusBar = False
                                                    Debug.WriteLine("状态栏已清除")
                                                End If
                                            Catch ex As Exception
                                                Debug.WriteLine($"清除状态栏失败: {ex.Message}")
                                            End Try
                                        End Sub)
        Catch ex As Exception
            Debug.WriteLine($"队列清除状态栏操作失败: {ex.Message}")
        End Try
    End Sub

    ' 同步处理LLM请求 - 移除特殊字符
    Private Function ProcessLLMRequestSync(prompt As String, model As String, systemPrompt As String,
                                         temperature As Double, maxTokens As Integer) As String
        Try
            ' 验证输入
            If String.IsNullOrEmpty(prompt) Then
                Return "错误: 提示词不能为空"
            End If

            EnsureConfigLoaded()
            Dim apiKey As String = ConfigSettings.ApiKey
            Dim apiUrl As String = ConfigSettings.ApiUrl

            If String.IsNullOrEmpty(apiKey) Then
                Return "错误: 未配置API密钥"
            End If

            If String.IsNullOrEmpty(apiUrl) Then
                Return "错误: 未配置API URL"
            End If

            ' 使用指定的模型或默认模型
            Dim useModel As String = If(String.IsNullOrEmpty(model), GetDefaultModel(), model)

            ' 创建请求体
            Dim requestBody As String = LLMUtil.CreateLlmRequestBody(prompt, useModel, systemPrompt, temperature, maxTokens)
            Debug.WriteLine($"请求体: {requestBody}")

            ' 使用同步HTTP请求
            Dim response As String = LLMUtil.SendHttpRequestSync(apiUrl, apiKey, requestBody)
            Debug.WriteLine($"收到HTTP响应: {response}")

            ' 如果响应为空或是错误信息，直接返回
            If String.IsNullOrEmpty(response) Then
                Return "错误: API未返回响应"
            End If

            If response.StartsWith("错误:") Then
                Return response
            End If

            ' 尝试解析JSON响应
            Dim parsedResponse As JObject = JObject.Parse(response)
            Dim cellValue As String = parsedResponse("choices")(0)("message")("content").ToString()
            Debug.WriteLine($"解析的响应内容: {cellValue}")

            ' 缓存结果
            Dim cacheKey As String = $"ALLM|{prompt}|{model}|{systemPrompt}|{temperature}|{maxTokens}"
            FunctionCache(cacheKey) = cellValue

            Return cellValue

        Catch ex As Exception
            Debug.WriteLine($"ProcessLLMRequestSync异常: {ex.Message}")
            Return $"错误: {ex.Message}"
        End Try
    End Function
    ' 缓存实现
    Private FunctionCache As New Dictionary(Of String, String)

    ' 获取默认模型
    Private Function GetDefaultModel() As String
        Try
            ' 从配置管理器获取默认模型
            Dim model As String = ConfigSettings.ModelName
            Return model
        Catch ex As Exception
            Return "gpt-3.5-turbo"
        End Try
    End Function
End Module