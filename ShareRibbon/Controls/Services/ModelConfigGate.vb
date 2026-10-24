' ShareRibbon\Controls\Services\ModelConfigGate.vb
' E3-03: 未配置模型时的面板内 CTA。阻断只发生在共享漏斗，文案与跳转集中在这里，
' 避免聊天、Agent、Ribbon 三条入口各写一套"请先配置"提示。

Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json

''' <summary>
''' 一次“配置缺失”阻断的可渲染描述。
''' </summary>
Public Class ModelConfigPrompt
    Public Property Reason As String
    Public Property Title As String
    Public Property Message As String
    Public Property ActionLabel As String
    Public Property BlockedAction As String

    Public Function ToJson() As String
        Dim payload As New JObject()
        payload("reason") = If(Reason, "")
        payload("title") = If(Title, "")
        payload("message") = If(Message, "")
        payload("actionLabel") = If(ActionLabel, "")
        payload("blockedAction") = If(BlockedAction, "")
        Return payload.ToString(Formatting.None)
    End Function
End Class

Public Class ModelConfigGate
    Public Const ActionChat As String = "chat"
    Public Const ActionAgent As String = "agent"

    Private Shared Function IsConfigFailure(failure As ChatSendValidationFailure) As Boolean
        Return failure = ChatSendValidationFailure.MissingApiKey OrElse
               failure = ChatSendValidationFailure.MissingApiUrl
    End Function

    Private Shared ReadOnly Validator As New ChatSendValidator()

    ''' <summary>
    ''' 配置可用时返回 Nothing：调用方据此决定是放行还是只给 CTA。
    ''' </summary>
    Public Function Evaluate(apiUrl As String, apiKey As String, blockedAction As String) As ModelConfigPrompt
        Dim result = Validator.ValidateConfig(apiUrl, apiKey)
        If Not IsConfigFailure(result.Failure) Then Return Nothing
        Return Build(result.Failure, blockedAction)
    End Function

    ''' <summary>
    ''' 已经把 ChatSendValidationResult 算出来的调用方（Send）用这个重载，避免重复校验。
    ''' </summary>
    Public Function Evaluate(result As ChatSendValidationResult, blockedAction As String) As ModelConfigPrompt
        If result Is Nothing OrElse Not IsConfigFailure(result.Failure) Then Return Nothing
        Return Build(result.Failure, blockedAction)
    End Function

    Private Shared Function Build(failure As ChatSendValidationFailure, blockedAction As String) As ModelConfigPrompt
        Dim action = If(blockedAction, ActionChat)
        Dim stopped = If(action = ActionAgent,
                         "任务没有开始执行，文档不会被修改。",
                         "消息没有发送，文档不会被修改。")
        Dim detail As String
        Select Case failure
            Case ChatSendValidationFailure.MissingApiKey
                detail = "还没有可用的 API Key。"
            Case Else
                detail = "还没有可用的接口地址（API 地址）。"
        End Select

        Return New ModelConfigPrompt With {
            .Reason = If(failure = ChatSendValidationFailure.MissingApiKey, "missing_api_key", "missing_api_url"),
            .Title = "需要先配置大模型",
            .Message = detail & stopped & "点击下方按钮填写接口地址和 Key，或选择本地模型（Ollama / LM Studio）。",
            .ActionLabel = "配置 API",
            .BlockedAction = action
        }
    End Function

    Public Function BuildShowScript(prompt As ModelConfigPrompt) As String
        If prompt Is Nothing Then Return ""
        Return "showModelConfigRequired(" & prompt.ToJson() & ");"
    End Function

    Public Shared Function BuildHideScript() As String
        Return "hideModelConfigRequired();"
    End Function
End Class
