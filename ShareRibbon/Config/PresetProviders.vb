Imports ShareRibbon.ConfigManager

''' <summary>
''' 预置服务商配置
''' 包含云端服务商和本地模型服务的默认配置
''' </summary>
Public Class PresetProviders

    ''' <summary>
    ''' 获取云端服务商预置配置
    ''' </summary>
    Public Shared Function GetCloudProviders() As List(Of ConfigItem)
        Dim providers As New List(Of ConfigItem)()

        ' 深度求索 (DeepSeek)
        Dim deepseek As New ConfigItem()
        deepseek.pltform = "深度求索 (DeepSeek)"
        deepseek.url = "https://api.deepseek.com/chat/completions"
        deepseek.registerUrl = "https://platform.deepseek.com/api_keys"
        deepseek.providerType = ProviderType.Cloud
        deepseek.isPreset = True
        deepseek.key = ""
        deepseek.selected = True
        deepseek.translateSelected = True
        deepseek.model = New List(Of ConfigItemModel)()
        deepseek.model.Add(CreateModel("deepseek-chat", "deepseek-chat [MCP]", True, True, False))
        deepseek.model.Add(CreateModel("deepseek-reasoner", "deepseek-reasoner [推理]", False, False, True))
        providers.Add(deepseek)

        ' 硅基流动 (SiliconFlow)
        Dim siliconflow As New ConfigItem()
        siliconflow.pltform = "硅基流动 (SiliconFlow)"
        siliconflow.url = "https://api.siliconflow.cn/v1/chat/completions"
        siliconflow.registerUrl = "https://cloud.siliconflow.cn/i/PGhr3knx"
        siliconflow.providerType = ProviderType.Cloud
        siliconflow.isPreset = True
        siliconflow.key = ""
        siliconflow.selected = False
        siliconflow.model = New List(Of ConfigItemModel)()
        siliconflow.model.Add(CreateModel("deepseek-ai/DeepSeek-V3", "DeepSeek-V3", True, False, False))
        siliconflow.model.Add(CreateModel("deepseek-ai/DeepSeek-R1", "DeepSeek-R1 [推理]", False, False, True))
        siliconflow.model.Add(CreateModel("Qwen/Qwen2.5-Coder-32B-Instruct", "Qwen2.5-Coder-32B", False, False, False))
        siliconflow.model.Add(CreateModel("Pro/deepseek-ai/DeepSeek-R1", "DeepSeek-R1-Pro [推理]", False, False, True))
        providers.Add(siliconflow)

        ' 阿里云百炼 (Qwen)
        Dim aliyun As New ConfigItem()
        aliyun.pltform = "阿里云百炼 (Qwen)"
        aliyun.url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
        aliyun.registerUrl = "https://help.aliyun.com/zh/model-studio/getting-started/first-api-call-to-qwen"
        aliyun.providerType = ProviderType.Cloud
        aliyun.isPreset = True
        aliyun.key = ""
        aliyun.selected = False
        aliyun.model = New List(Of ConfigItemModel)()
        aliyun.model.Add(CreateModel("qwen-coder-plus", "qwen-coder-plus", True, False, False))
        aliyun.model.Add(CreateModel("qwen-max", "qwen-max", False, False, False))
        aliyun.model.Add(CreateModel("qwen-plus", "qwen-plus", False, False, False))
        aliyun.model.Add(CreateModel("qwq-32b-preview", "qwq-32b-preview [推理]", False, False, True))
        providers.Add(aliyun)

        ' 百度千帆
        Dim baidu As New ConfigItem()
        baidu.pltform = "百度千帆"
        baidu.url = "https://qianfan.baidubce.com/v2/chat/completions"
        baidu.registerUrl = "https://console.bce.baidu.com/qianfan/ais/console/applicationConsole/application"
        baidu.providerType = ProviderType.Cloud
        baidu.isPreset = True
        baidu.key = ""
        baidu.selected = False
        baidu.model = New List(Of ConfigItemModel)()
        baidu.model.Add(CreateModel("deepseek-v3", "deepseek-v3", True, False, False))
        baidu.model.Add(CreateModel("deepseek-r1", "deepseek-r1 [推理]", False, False, True))
        baidu.model.Add(CreateModel("ernie-4.0-8k", "ernie-4.0-8k", False, False, False))
        providers.Add(baidu)

        ' 智谱清言 (GLM)
        Dim zhipu As New ConfigItem()
        zhipu.pltform = "智谱清言 (GLM)"
        zhipu.url = "https://open.bigmodel.cn/api/paas/v4/chat/completions"
        zhipu.registerUrl = "https://www.bigmodel.cn/invite?icode=pRDixwFdhElsS8rrQ7JbplwpqjqOwPB5EXW6OL4DgqY%3D"
        '    https://open.bigmodel.cn/usercenter/apikeys
        zhipu.providerType = ProviderType.Cloud
        zhipu.isPreset = True
        zhipu.key = ""
        zhipu.selected = False
        zhipu.model = New List(Of ConfigItemModel)()
        zhipu.model.Add(CreateModel("glm-4-plus", "glm-4-plus [MCP]", True, True, False))
        zhipu.model.Add(CreateModel("glm-4-flash", "glm-4-flash", False, False, False))
        zhipu.model.Add(CreateModel("glm-4-air", "glm-4-air", False, False, False))
        zhipu.model.Add(CreateModel("glm-4-long", "glm-4-long", False, False, False))
        providers.Add(zhipu)

        ' 腾讯混元 (Hunyuan)
        Dim tencent As New ConfigItem()
        tencent.pltform = "腾讯混元 (Hunyuan)"
        tencent.url = "https://api.hunyuan.cloud.tencent.com/v1/chat/completions"
        tencent.registerUrl = "https://cloud.tencent.com/product/hunyuan"
        tencent.providerType = ProviderType.Cloud
        tencent.isPreset = True
        tencent.key = ""
        tencent.selected = False
        tencent.model = New List(Of ConfigItemModel)()
        tencent.model.Add(CreateModel("hunyuan-turbo", "hunyuan-turbo", True, False, False))
        tencent.model.Add(CreateModel("hunyuan-pro", "hunyuan-pro", False, False, False))
        tencent.model.Add(CreateModel("hunyuan-standard", "hunyuan-standard", False, False, False))
        providers.Add(tencent)

        ' OpenRouter
        Dim openrouter As New ConfigItem()
        openrouter.pltform = "OpenRouter"
        openrouter.url = "https://openrouter.ai/api/v1/chat/completions"
        openrouter.registerUrl = "https://openrouter.ai/keys"
        openrouter.providerType = ProviderType.Cloud
        openrouter.isPreset = True
        openrouter.key = ""
        openrouter.selected = False
        openrouter.model = New List(Of ConfigItemModel)()
        openrouter.model.Add(CreateModel("deepseek/deepseek-r1", "deepseek-r1 [推理][MCP]", True, True, True))
        openrouter.model.Add(CreateModel("deepseek/deepseek-chat", "deepseek-chat [MCP]", False, True, False))
        openrouter.model.Add(CreateModel("anthropic/claude-3.5-sonnet", "claude-3.5-sonnet [MCP]", False, True, False))
        openrouter.model.Add(CreateModel("openai/gpt-4o", "gpt-4o [MCP]", False, True, False))
        openrouter.model.Add(CreateModel("google/gemini-2.0-flash-exp:free", "gemini-2.0-flash [免费]", False, False, False))
        providers.Add(openrouter)

        ' Kimi (Moonshot)
        Dim kimi As New ConfigItem()
        kimi.pltform = "Kimi (Moonshot)"
        kimi.url = "https://api.moonshot.cn/v1/chat/completions"
        kimi.registerUrl = "https://platform.moonshot.cn/console/api-keys"
        kimi.providerType = ProviderType.Cloud
        kimi.isPreset = True
        kimi.key = ""
        kimi.selected = False
        kimi.model = New List(Of ConfigItemModel)()
        kimi.model.Add(CreateModel("moonshot-v1-8k", "moonshot-v1-8k", True, False, False))
        kimi.model.Add(CreateModel("moonshot-v1-32k", "moonshot-v1-32k", False, False, False))
        kimi.model.Add(CreateModel("moonshot-v1-128k", "moonshot-v1-128k", False, False, False))
        providers.Add(kimi)

        ' Google Gemini
        Dim gemini As New ConfigItem()
        gemini.pltform = "Google Gemini"
        gemini.url = "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions"
        gemini.registerUrl = "https://aistudio.google.com/apikey"
        gemini.providerType = ProviderType.Cloud
        gemini.isPreset = True
        gemini.key = ""
        gemini.selected = False
        gemini.model = New List(Of ConfigItemModel)()
        gemini.model.Add(CreateModel("gemini-2.0-flash-exp", "gemini-2.0-flash [推理]", True, False, True))
        gemini.model.Add(CreateModel("gemini-1.5-pro", "gemini-1.5-pro", False, False, False))
        gemini.model.Add(CreateModel("gemini-1.5-flash", "gemini-1.5-flash", False, False, False))
        providers.Add(gemini)

        ' OpenAI ChatGPT
        Dim openai As New ConfigItem()
        openai.pltform = "OpenAI ChatGPT"
        openai.url = "https://api.openai.com/v1/chat/completions"
        openai.registerUrl = "https://platform.openai.com/api-keys"
        openai.providerType = ProviderType.Cloud
        openai.isPreset = True
        openai.key = ""
        openai.selected = False
        openai.model = New List(Of ConfigItemModel)()
        openai.model.Add(CreateModel("gpt-4o", "gpt-4o [MCP]", True, True, False))
        openai.model.Add(CreateModel("gpt-4o-mini", "gpt-4o-mini [MCP]", False, True, False))
        openai.model.Add(CreateModel("gpt-4-turbo", "gpt-4-turbo", False, False, False))
        openai.model.Add(CreateModel("o1", "o1 [推理]", False, False, True))
        openai.model.Add(CreateModel("o1-mini", "o1-mini [推理]", False, False, True))
        providers.Add(openai)

        ' Grok (xAI)
        Dim grok As New ConfigItem()
        grok.pltform = "Grok (xAI)"
        grok.url = "https://api.x.ai/v1/chat/completions"
        grok.registerUrl = "https://console.x.ai/"
        grok.providerType = ProviderType.Cloud
        grok.isPreset = True
        grok.key = ""
        grok.selected = False
        grok.model = New List(Of ConfigItemModel)()
        grok.model.Add(CreateModel("grok-2-1212", "grok-2-1212", True, False, False))
        grok.model.Add(CreateModel("grok-beta", "grok-beta", False, False, False))
        providers.Add(grok)

        ' Anthropic Claude
        Dim anthropic As New ConfigItem()
        anthropic.pltform = "Anthropic Claude"
        anthropic.url = "https://api.anthropic.com/v1/messages"
        anthropic.registerUrl = "https://console.anthropic.com/settings/keys"
        anthropic.providerType = ProviderType.Cloud
        anthropic.isPreset = True
        anthropic.key = ""
        anthropic.selected = False
        anthropic.model = New List(Of ConfigItemModel)()
        anthropic.model.Add(CreateModel("claude-3-5-sonnet-20241022", "claude-3.5-sonnet [MCP]", True, True, False))
        anthropic.model.Add(CreateModel("claude-3-5-haiku-20241022", "claude-3.5-haiku", False, False, False))
        anthropic.model.Add(CreateModel("claude-3-opus-20240229", "claude-3-opus", False, False, False))
        providers.Add(anthropic)

        Return providers
    End Function

    ''' <summary>
    ''' 获取本地模型服务预置配置
    ''' </summary>
    Public Shared Function GetLocalProviders() As List(Of ConfigItem)
        Dim providers As New List(Of ConfigItem)()

        ' Ollama
        Dim ollama As New ConfigItem()
        ollama.pltform = "Ollama"
        ollama.url = "http://localhost:11434/v1/chat/completions"
        ollama.providerType = ProviderType.Local
        ollama.isPreset = True
        ollama.key = "ollama"
        ollama.defaultApiKey = "ollama"
        ollama.selected = False
        ollama.model = New List(Of ConfigItemModel)()
        ollama.model.Add(CreateModel("deepseek-r1:1.5b", "deepseek-r1:1.5b [推理]", True, False, True))
        ollama.model.Add(CreateModel("deepseek-r1:7b", "deepseek-r1:7b [推理]", False, False, True))
        ollama.model.Add(CreateModel("deepseek-r1:14b", "deepseek-r1:14b [推理]", False, False, True))
        ollama.model.Add(CreateModel("qwen2.5-coder:7b", "qwen2.5-coder:7b", False, False, False))
        ollama.model.Add(CreateModel("llama3.2:latest", "llama3.2:latest", False, False, False))
        providers.Add(ollama)

        ' vLLM
        Dim vllm As New ConfigItem()
        vllm.pltform = "vLLM"
        vllm.url = "http://localhost:8000/v1/chat/completions"
        vllm.providerType = ProviderType.Local
        vllm.isPreset = True
        vllm.key = ""
        vllm.defaultApiKey = "vllm"
        vllm.selected = False
        vllm.model = New List(Of ConfigItemModel)()
        providers.Add(vllm)

        ' LM Studio
        Dim lmstudio As New ConfigItem()
        lmstudio.pltform = "LM Studio"
        lmstudio.url = "http://localhost:1234/v1/chat/completions"
        lmstudio.providerType = ProviderType.Local
        lmstudio.isPreset = True
        lmstudio.key = ""
        lmstudio.defaultApiKey = "lm-studio"
        lmstudio.selected = False
        lmstudio.model = New List(Of ConfigItemModel)()
        providers.Add(lmstudio)

        ' RWKV
        Dim rwkv As New ConfigItem()
        rwkv.pltform = "RWKV"
        rwkv.url = "http://localhost:8080/v1/chat/completions"
        rwkv.providerType = ProviderType.Local
        rwkv.isPreset = True
        rwkv.key = ""
        rwkv.defaultApiKey = "rwkv"
        rwkv.selected = False
        rwkv.model = New List(Of ConfigItemModel)()
        providers.Add(rwkv)

        ' LMDeploy
        Dim lmdeploy As New ConfigItem()
        lmdeploy.pltform = "LMDeploy"
        lmdeploy.url = "http://localhost:23333/v1/chat/completions"
        lmdeploy.providerType = ProviderType.Local
        lmdeploy.isPreset = True
        lmdeploy.key = ""
        lmdeploy.defaultApiKey = "lmdeploy"
        lmdeploy.selected = False
        lmdeploy.model = New List(Of ConfigItemModel)()
        providers.Add(lmdeploy)

        Return providers
    End Function

    ''' <summary>
    ''' 创建模型配置项的辅助方法
    ''' </summary>
    Private Shared Function CreateModel(modelName As String, displayName As String, selected As Boolean, mcpable As Boolean, isReasoning As Boolean) As ConfigItemModel
        Dim model As New ConfigItemModel()
        model.modelName = modelName
        model.displayName = displayName
        model.selected = selected
        model.mcpable = mcpable
        model.isReasoningModel = isReasoning
        Return model
    End Function

    ''' <summary>
    ''' 获取所有预置配置
    ''' </summary>
    Public Shared Function GetAllPresetProviders() As List(Of ConfigItem)
        Dim allProviders As New List(Of ConfigItem)()
        allProviders.AddRange(GetCloudProviders())
        allProviders.AddRange(GetLocalProviders())
        Return allProviders
    End Function

End Class
