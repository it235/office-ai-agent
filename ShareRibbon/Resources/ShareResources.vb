Public Class SharedResources
    Private Shared ReadOnly _resourceManager As New System.Resources.ResourceManager("ShareRibbon.Resources", GetType(SharedResources).Assembly)

    Public Shared ReadOnly Property AiApiConfig() As System.Drawing.Image
        Get
            Return CType(_resourceManager.GetObject("aiapiconfig"), System.Drawing.Image)
        End Get
    End Property

    Public Shared ReadOnly Property Magic() As System.Drawing.Image
        Get
            Return CType(_resourceManager.GetObject("magic"), System.Drawing.Image)
        End Get
    End Property

    Public Shared ReadOnly Property Send32() As System.Drawing.Image
        Get
            Return CType(_resourceManager.GetObject("send32"), System.Drawing.Image)
        End Get
    End Property

    Public Shared ReadOnly Property Mcp1() As System.Drawing.Image
        Get
            Return CType(_resourceManager.GetObject("mcp1"), System.Drawing.Image)
        End Get
    End Property
    Public Shared ReadOnly Property Chat() As System.Drawing.Image
        Get
            Return CType(_resourceManager.GetObject("chat"), System.Drawing.Image)
        End Get
    End Property

    Public Shared ReadOnly Property About() As System.Drawing.Image
        Get
            Return CType(_resourceManager.GetObject("about"), System.Drawing.Image)
        End Get
    End Property

    Public Shared ReadOnly Property Clear() As System.Drawing.Image
        Get
            Return CType(_resourceManager.GetObject("clear"), System.Drawing.Image)
        End Get
    End Property
    Public Shared ReadOnly Property Deepseek() As System.Drawing.Image
        Get
            Return CType(_resourceManager.GetObject("deepseek"), System.Drawing.Image)
        End Get
    End Property
    Public Shared ReadOnly Property Doubao() As System.Drawing.Image
        Get
            Return CType(_resourceManager.GetObject("doubao_avatar"), System.Drawing.Image)
        End Get
    End Property
    Public Shared ReadOnly Property Wait() As System.Drawing.Image
        Get
            Return CType(_resourceManager.GetObject("wait"), System.Drawing.Image)
        End Get
    End Property
    Public Shared ReadOnly Property Help() As System.Drawing.Image
        Get
            Return CType(_resourceManager.GetObject("help"), System.Drawing.Image)
        End Get
    End Property
    Public Shared ReadOnly Property Audit() As System.Drawing.Image
        Get
            Return CType(_resourceManager.GetObject("audit"), System.Drawing.Image)
        End Get
    End Property
    Public Shared ReadOnly Property Papers() As System.Drawing.Image
        Get
            Return CType(_resourceManager.GetObject("papers"), System.Drawing.Image)
        End Get
    End Property
    Public Shared ReadOnly Property Translate() As System.Drawing.Image
        Get
            Return CType(_resourceManager.GetObject("translate"), System.Drawing.Image)
        End Get
    End Property
    Public Shared ReadOnly Property Aiwrite() As System.Drawing.Image
        Get
            Return CType(_resourceManager.GetObject("aiwrite"), System.Drawing.Image)
        End Get
    End Property

    Public Shared ReadOnly Property autocomplete() As System.Drawing.Image
        Get
            Return CType(_resourceManager.GetObject("autocomplete1"), System.Drawing.Image)
        End Get
    End Property

    Public Shared ReadOnly Property promptconfig() As System.Drawing.Image
        Get
            Return CType(_resourceManager.GetObject("promptconfig"), System.Drawing.Image)
        End Get
    End Property
End Class