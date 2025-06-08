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
End Class