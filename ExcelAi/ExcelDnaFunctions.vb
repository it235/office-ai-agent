Imports System.Diagnostics
Imports System.Threading.Tasks
Imports ExcelDna.Integration
Imports Newtonsoft.Json.Linq
Imports ShareRibbon

Public Module ExcelDnaFunctions

    ' ���ù�����ʵ��
    Private configManager As ConfigManager = Nothing

    ' �޸�AutoOpen���� - �Ƴ�MenuName��MenuText����
    <ExcelDna.Integration.ExcelCommand()>
    Public Sub AutoOpen()
        Try
            Debug.WriteLine("=== Excel-DNA AutoOpen ��ʼִ�� ===")
            Console.WriteLine("=== Excel-DNA AutoOpen ��ʼִ�� ===")

            ' ��Excel-DNA���г�ʼ��GlobalStatusStrip
            Try
                Dim excelApp As Object = ExcelDna.Integration.ExcelDnaUtil.Application
                If excelApp IsNot Nothing Then
                    GlobalStatusStripAll.InitializeApplication(excelApp)
                    Debug.WriteLine("Excel-DNA����GlobalStatusStrip��ʼ���ɹ�")

                    ' ����״̬���Ƿ���
                    Try
                        excelApp.StatusBar = "Excel-DNA�Ѽ���"
                        Debug.WriteLine("Excel-DNA״̬�����Գɹ�")
                        Threading.Thread.Sleep(2000)
                        excelApp.StatusBar = False
                    Catch statusEx As Exception
                        Debug.WriteLine($"Excel-DNA״̬������ʧ��: {statusEx.Message}")
                    End Try
                Else
                    Debug.WriteLine("Excel-DNA�����޷���ȡExcelӦ�ó������")
                End If
            Catch ex As Exception
                Debug.WriteLine($"Excel-DNA����GlobalStatusStrip��ʼ��ʧ��: {ex.Message}")
            End Try

            ' ��ʼ�����ù�����
            Try
                configManager = New ConfigManager()
                configManager.LoadConfig()
                Debug.WriteLine($"Excel-DNA �����Ѽ��� - API URL: {ConfigSettings.ApiUrl}")
                Debug.WriteLine($"Excel-DNA �����Ѽ��� - Model: {ConfigSettings.ModelName}")
            Catch ex As Exception
                Debug.WriteLine($"Excel-DNA ���ü���ʧ��: {ex.Message}")
            End Try

            ' �������
            FunctionCache.Clear()

            ' ע��δ�����쳣�������
            ExcelDna.Integration.ExcelIntegration.RegisterUnhandledExceptionHandler(
            Function(ex) $"����: {ex.Message}")

            Debug.WriteLine("=== Excel-DNA AutoOpen ִ����� ===")
            Console.WriteLine("=== Excel-DNA AutoOpen ִ����� ===")

        Catch ex As Exception
            Debug.WriteLine($"Excel-DNA AutoOpen ִ��ʧ��: {ex.Message}")
            Console.WriteLine($"Excel-DNA AutoOpen ִ��ʧ��: {ex.Message}")
        End Try
    End Sub


    ' ȷ�������Ѽ���
    Private Sub EnsureConfigLoaded()
        If configManager Is Nothing Then
            configManager = New ConfigManager()
            configManager.LoadConfig()
        End If
    End Sub

    ' ���� LLM ����
    <ExcelFunction(Description:="ʹ��AIģ�������ı�",
                  Category:="Excel AI ����",
                  Name:="ELLM",
                  IsVolatile:=False,
                  IsThreadSafe:=True)>
    Public Function ELLM(
        <ExcelArgument(Description:="��ʾ�ʻ�����")> prompt As String) As Object

        ' ʹ�û�������ظ�����
        Dim cacheKey As String = $"ELLM|{prompt}"
        If FunctionCache.ContainsKey(cacheKey) Then
            Return FunctionCache(cacheKey)
        End If

        Try
            ' ���������溯��
            Dim result As String = ADLLM(prompt, "", "", 0.7, 1000)

            ' ������
            If Not result.StartsWith("����:") Then
                FunctionCache(cacheKey) = result
            End If

            Return result
        Catch ex As Exception
            Return $"����: {ex.Message}"
        End Try
    End Function

    ' �߼� LLM ����
    <ExcelFunction(Description:="ʹ��AIģ�������ı�(�߼���)",
                  Category:="Excel AI ����",
                  Name:="ADLLM",
                  IsVolatile:=False,
                  IsThreadSafe:=True)>
    Public Function ADLLM(
        <ExcelArgument(Description:="��ʾ�ʻ�����")> prompt As String,
        <ExcelArgument(Description:="��ѡ: ģ������")> Optional model As String = "",
        <ExcelArgument(Description:="��ѡ: ϵͳ��ʾ��")> Optional systemPrompt As String = "",
        <ExcelArgument(Description:="��ѡ: �¶Ȳ��� (0.0-1.0)")> Optional temperature As Double = 0.7,
        <ExcelArgument(Description:="��ѡ: �������������")> Optional maxTokens As Integer = 1000) As Object

        ' ʹ�û�������ظ�����
        Dim cacheKey As String = $"ADLLM|{prompt}|{model}|{systemPrompt}|{temperature}|{maxTokens}"
        If FunctionCache.ContainsKey(cacheKey) Then
            Return FunctionCache(cacheKey)
        End If

        Try
            ' ��֤����
            If String.IsNullOrEmpty(prompt) Then
                Return "����: ��ʾ�ʲ���Ϊ��"
            End If

            ' ȷ�������Ѽ���
            EnsureConfigLoaded()
            ' ʹ��Ĭ��ֵ�����δ�ṩ��
            Dim apiKey As String = ConfigSettings.ApiKey
            Dim apiUrl As String = ConfigSettings.ApiUrl

            If String.IsNullOrEmpty(apiKey) Then
                Return "����: δ����API��Կ"
            End If

            If String.IsNullOrEmpty(apiUrl) Then
                Return "����: δ����API URL"
            End If

            ' ʹ��ָ����ģ�ͻ�Ĭ��ģ��
            Dim useModel As String = If(String.IsNullOrEmpty(model), GetDefaultModel(), model)

            GlobalStatusStripAll.ShowWarning($"����ͬ�������ģ���У������ĵȴ�")

            ' ����������
            Dim requestBody As String = LLMUtil.CreateLlmRequestBody(prompt, useModel, systemPrompt, temperature, maxTokens)

            ' ����API����ȡ���
            Dim response As String = LLMUtil.SendHttpRequest(apiUrl, apiKey, requestBody).Result

            ' �����ӦΪ�գ����ش�����Ϣ
            If String.IsNullOrEmpty(response) Then
                Return "����: APIδ������Ӧ"
            End If

            Dim parsedResponse As JObject = JObject.Parse(response)
            Dim cellValue As String = parsedResponse("choices")(0)("message")("content").ToString()

            ' ������
            FunctionCache(cacheKey) = cellValue

            Return cellValue
        Catch ex As Exception
            Return $"����: {ex.Message}"
        End Try
    End Function

    ' �첽 LLM ���� - �޸��汾��ȥ��δ����Ľӿ�
    <ExcelFunction(Description:="�첽����AIģ�������ı�",
                  Category:="Excel AI ����",
                  Name:="ALLM",
                  IsVolatile:=False,
                  IsThreadSafe:=True)>
    Public Function ALLM(
        <ExcelArgument(Description:="��ʾ�ʻ�����")> prompt As String,
        <ExcelArgument(Description:="��ѡ: ģ������")> Optional model As String = "",
        <ExcelArgument(Description:="��ѡ: ϵͳ��ʾ��")> Optional systemPrompt As String = "",
        <ExcelArgument(Description:="��ѡ: �¶Ȳ��� (0.0-1.0)")> Optional temperature As Double = 0.7,
        <ExcelArgument(Description:="��ѡ: �������������")> Optional maxTokens As Integer = 1000) As Object

        If temperature = 0.0 Then temperature = 0.7
        If maxTokens = 0 Then maxTokens = 1000

        ' ʹ��Excel-DNA���е��̰߳�ȫ��������״̬��
        Try
            Dim displayPrompt As String = If(prompt.Length > 25, prompt.Substring(0, 25) + "...", prompt)
            SetExcelStatusBarDirectly($"����˼����{displayPrompt}��...")
        Catch ex As Exception
            Debug.WriteLine($"��ʾ״̬��ʾʧ��: {ex.Message}")
        End Try

        ' ��黺�棬����л���ֱ�ӷ���
        Dim cacheKey As String = $"ALLM|{prompt}|{model}|{systemPrompt}|{temperature}|{maxTokens}"
        If FunctionCache.ContainsKey(cacheKey) Then
            Return FunctionCache(cacheKey)
        End If


        ' ʹ�� ExcelAsyncUtil.Run ִ���첽����
        Return ExcelAsyncUtil.Run("ALLM", New Object() {prompt, model, systemPrompt, temperature, maxTokens},
            Function()
                Try
                    Dim result As String = ProcessLLMRequestSync(prompt, model, systemPrompt, temperature, maxTokens)

                    ' ���״̬��
                    ClearExcelStatusBar()
                    Return result
                Catch ex As Exception
                    ClearExcelStatusBar()
                    Return $"����: {ex.Message}"
                End Try
            End Function)
    End Function

    ' ר������Excel-DNA���̰߳�ȫ״̬������
    Private Sub SetExcelStatusBarDirectly(message As String)
        Try
            Debug.WriteLine($"���ڳ��԰�ȫ����״̬��: {message}")

            ' ʹ��Excel-DNA��QueueAsMacroȷ����Excel���߳���ִ��
            ExcelAsyncUtil.QueueAsMacro(Sub()
                                            Try
                                                Dim excelApp As Object = ExcelDna.Integration.ExcelDnaUtil.Application
                                                If excelApp IsNot Nothing Then
                                                    excelApp.StatusBar = message
                                                    Debug.WriteLine($"״̬�����óɹ�: {message}")
                                                Else
                                                    Debug.WriteLine("�޷���ȡExcelӦ�ó������")
                                                End If
                                            Catch innerEx As Exception
                                                Debug.WriteLine($"��Excel���߳�������״̬��ʧ��: {innerEx.Message}")
                                            End Try
                                        End Sub)
        Catch ex As Exception
            Debug.WriteLine($"����״̬������ʧ��: {ex.Message}")
        End Try
    End Sub

    ' ��ProcessLLMRequestSync��ɺ����״̬��
    Private Sub ClearExcelStatusBar()
        Try
            ExcelAsyncUtil.QueueAsMacro(Sub()
                                            Try
                                                Dim excelApp As Object = ExcelDna.Integration.ExcelDnaUtil.Application
                                                If excelApp IsNot Nothing Then
                                                    excelApp.StatusBar = False
                                                    Debug.WriteLine("״̬�������")
                                                End If
                                            Catch ex As Exception
                                                Debug.WriteLine($"���״̬��ʧ��: {ex.Message}")
                                            End Try
                                        End Sub)
        Catch ex As Exception
            Debug.WriteLine($"�������״̬������ʧ��: {ex.Message}")
        End Try
    End Sub

    ' ͬ������LLM���� - �Ƴ������ַ�
    Private Function ProcessLLMRequestSync(prompt As String, model As String, systemPrompt As String,
                                         temperature As Double, maxTokens As Integer) As String
        Try
            ' ��֤����
            If String.IsNullOrEmpty(prompt) Then
                Return "����: ��ʾ�ʲ���Ϊ��"
            End If

            EnsureConfigLoaded()
            Dim apiKey As String = ConfigSettings.ApiKey
            Dim apiUrl As String = ConfigSettings.ApiUrl

            If String.IsNullOrEmpty(apiKey) Then
                Return "����: δ����API��Կ"
            End If

            If String.IsNullOrEmpty(apiUrl) Then
                Return "����: δ����API URL"
            End If

            ' ʹ��ָ����ģ�ͻ�Ĭ��ģ��
            Dim useModel As String = If(String.IsNullOrEmpty(model), GetDefaultModel(), model)

            ' ����������
            Dim requestBody As String = LLMUtil.CreateLlmRequestBody(prompt, useModel, systemPrompt, temperature, maxTokens)
            Debug.WriteLine($"������: {requestBody}")

            ' ʹ��ͬ��HTTP����
            Dim response As String = LLMUtil.SendHttpRequestSync(apiUrl, apiKey, requestBody)
            Debug.WriteLine($"�յ�HTTP��Ӧ: {response}")

            ' �����ӦΪ�ջ��Ǵ�����Ϣ��ֱ�ӷ���
            If String.IsNullOrEmpty(response) Then
                Return "����: APIδ������Ӧ"
            End If

            If response.StartsWith("����:") Then
                Return response
            End If

            ' ���Խ���JSON��Ӧ
            Dim parsedResponse As JObject = JObject.Parse(response)
            Dim cellValue As String = parsedResponse("choices")(0)("message")("content").ToString()
            Debug.WriteLine($"��������Ӧ����: {cellValue}")

            ' ������
            Dim cacheKey As String = $"ALLM|{prompt}|{model}|{systemPrompt}|{temperature}|{maxTokens}"
            FunctionCache(cacheKey) = cellValue

            Return cellValue

        Catch ex As Exception
            Debug.WriteLine($"ProcessLLMRequestSync�쳣: {ex.Message}")
            Return $"����: {ex.Message}"
        End Try
    End Function
    ' ����ʵ��
    Private FunctionCache As New Dictionary(Of String, String)

    ' ��ȡĬ��ģ��
    Private Function GetDefaultModel() As String
        Try
            ' �����ù�������ȡĬ��ģ��
            Dim model As String = ConfigSettings.ModelName
            Return model
        Catch ex As Exception
            Return "gpt-3.5-turbo"
        End Try
    End Function
End Module