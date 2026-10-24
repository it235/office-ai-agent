Public Enum ChatSendValidationFailure
    None = 0
    MissingApiKey
    MissingApiUrl
    MissingQuestion
End Enum

Public Class ChatSendValidationResult
    Public Property IsValid As Boolean
    Public Property Failure As ChatSendValidationFailure

    Public Shared Function Success() As ChatSendValidationResult
        Return New ChatSendValidationResult With {
            .IsValid = True,
            .Failure = ChatSendValidationFailure.None
        }
    End Function

    Public Shared Function Fail(failure As ChatSendValidationFailure) As ChatSendValidationResult
        Return New ChatSendValidationResult With {
            .IsValid = False,
            .Failure = failure
        }
    End Function
End Class

Public Class ChatSendValidator
    ''' <summary>
    ''' 只看模型配置，不看问题文本。聊天与 Agent 两条入口都必须先过这把门，
    ''' 否则未配置时会把请求送进主链、在半路才失败，用户只看到状态栏一行红字。
    ''' </summary>
    Public Function ValidateConfig(apiUrl As String, apiKey As String) As ChatSendValidationResult
        If String.IsNullOrWhiteSpace(apiKey) Then
            Return ChatSendValidationResult.Fail(ChatSendValidationFailure.MissingApiKey)
        End If

        If String.IsNullOrWhiteSpace(apiUrl) Then
            Return ChatSendValidationResult.Fail(ChatSendValidationFailure.MissingApiUrl)
        End If

        Return ChatSendValidationResult.Success()
    End Function

    Public Function Validate(apiUrl As String, apiKey As String, question As String) As ChatSendValidationResult
        Dim config = ValidateConfig(apiUrl, apiKey)
        If Not config.IsValid Then Return config

        If String.IsNullOrWhiteSpace(question) Then
            Return ChatSendValidationResult.Fail(ChatSendValidationFailure.MissingQuestion)
        End If

        Return ChatSendValidationResult.Success()
    End Function
End Class
