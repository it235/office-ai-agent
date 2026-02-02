''' <summary>
''' 模型服务商类型枚举
''' </summary>
Public Enum ProviderType
    ''' <summary>
    ''' 云端模型服务商 (如OpenAI、智谱清言等)
    ''' </summary>
    Cloud = 0

    ''' <summary>
    ''' 本地模型服务 (如Ollama、vLLM等)
    ''' </summary>
    Local = 1
End Enum
