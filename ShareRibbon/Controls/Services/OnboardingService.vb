' ShareRibbon\Controls\Services\OnboardingService.vb
' E3-01 首次使用引导：只做"隐私确认 -> 示例任务"这两件以前没有的事。
' 服务商、API Key、连接测试沿用 ConfigApiForm 那一套，这里不写第二份配置代码，也不碰 Harness/Loop。

Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon.Security

''' <summary>
''' 面板内引导向导的状态与动作。
''' </summary>
Public Class OnboardingService

    Private ReadOnly _executeScript As Func(Of String, Task)
    Private ReadOnly _getAppType As Func(Of String)
    Private ReadOnly _onConfigChanged As Action

    Public Sub New(executeScript As Func(Of String, Task), getAppType As Func(Of String), onConfigChanged As Action)
        _executeScript = executeScript
        _getAppType = getAppType
        _onConfigChanged = onConfigChanged
    End Sub

    ''' <summary>
    ''' 面板打开、以及配置窗口关闭后由前端拉状态；配置完整且隐私已确认时不会画向导，避免打扰老用户。
    ''' </summary>
    Public Sub HandleGetState()
        Try
            Send("renderOnboarding", BuildState())
        Catch ex As Exception
            AppLogger.Debug("OnboardingService", "构建引导状态失败", ex)
        End Try
    End Sub

    Public Sub HandleSetPrivacy(jsonDoc As JObject)
        Dim accepted = jsonDoc IsNot Nothing AndAlso jsonDoc("accepted") IsNot Nothing AndAlso
                       CStr(jsonDoc("accepted").ToString()).Equals("true", StringComparison.OrdinalIgnoreCase)
        Try
            SecureConfigManager.PrivacySettings.PrivacyPolicyAccepted = accepted
            Send("renderOnboarding", BuildState())
        Catch ex As Exception
            AppLogger.Debug("OnboardingService", "写入隐私确认失败", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 走完向导后刷新一次模型显示，让面板顶部立刻显示当前模型，而不是等下一次发送。
    ''' </summary>
    Public Sub HandleFinish()
        Try
            If _onConfigChanged IsNot Nothing Then _onConfigChanged()
        Catch ex As Exception
            AppLogger.Debug("OnboardingService", "引导收尾刷新失败", ex)
        End Try
    End Sub

    Private Function BuildState() As String
        Dim isValid = New ChatSendValidator().ValidateConfig(ConfigSettings.ApiUrl, ConfigSettings.ApiKey).IsValid
        Dim state As New JObject()
        state("needsSetup") = Not isValid
        state("privacyAccepted") = SecureConfigManager.PrivacySettings.PrivacyPolicyAccepted
        state("appType") = If(_getAppType(), "")
        state("step") = InitialStep()
        state("examples") = BuildExamples()
        Return state.ToString(Formatting.None)
    End Function

    ''' <summary>
    ''' 向导只剩两步，起点完全由配置状态决定：没确认过隐私就先确认，否则直接给示例。
    ''' </summary>
    Private Function InitialStep() As String
        If Not SecureConfigManager.PrivacySettings.PrivacyPolicyAccepted Then Return "privacy"
        Return "example"
    End Function

    Private Function BuildExamples() As JArray
        Dim appType = If(_getAppType(), "").Trim().ToLowerInvariant()
        Dim items As New List(Of String)
        Select Case appType
            Case "excel"
                items.Add("读取当前选区，找出异常值并说明原因")
                items.Add("为当前数据加一列同比增长率公式")
                items.Add("把选区整理成一张柱状图")
            Case "powerpoint"
                items.Add("按当前大纲生成 5 页汇报初稿")
                items.Add("统一全部标题的字体和字号")
                items.Add("检查哪一页文字过密并给出拆分建议")
            Case Else
                items.Add("把当前选区改写成更正式的书面表达")
                items.Add("检查全文的编号和标题层级是否连贯")
                items.Add("为当前文档生成一段摘要")
        End Select

        Dim arr As New JArray()
        For Each item In items
            arr.Add(item)
        Next
        Return arr
    End Function

    Private Async Function SendAsync(jsFunction As String, payloadJson As String) As Task
        If _executeScript Is Nothing Then Return
        Try
            Await _executeScript($"{jsFunction}({payloadJson});")
        Catch ex As Exception
            AppLogger.Debug("OnboardingService", $"推送 {jsFunction} 失败", ex)
        End Try
    End Function

    ''' <summary>
    ''' 同步处理器只需要"最终送达"，不等 WebView 回调。
    ''' </summary>
    Private Async Sub Send(jsFunction As String, payloadJson As String)
        Await SendAsync(jsFunction, payloadJson)
    End Sub

End Class
