' ShareRibbon\Controls\Services\OnboardingService.vb
' E3-01 首次使用引导：只做"隐私确认 -> 示例任务"这两件以前没有的事；E3-02 让每端第一条示例可以一键跑。
' 服务商、API Key、连接测试沿用 ConfigApiForm 那一套，这里不写第二份配置代码，也不碰 Harness/Loop。

Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

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
            PrivacyConsentGate.Apply(accepted)
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
        state("privacyAccepted") = PrivacyConsentGate.IsAccepted()
        ' 数据边界的三条说明由持有者下发，前端只渲染：文案在 JS 里再抄一份就等于有了第二份承诺。
        state("privacyBoundary") = JArray.FromObject(PrivacyConsentGate.BoundaryItems)
        state("appType") = If(_getAppType(), "")
        state("step") = InitialStep()
        state("examples") = BuildExamples()
        Return state.ToString(Formatting.None)
    End Function

    ''' <summary>
    ''' 向导只剩两步，起点完全由配置状态决定：没确认过隐私就先确认，否则直接给示例。
    ''' </summary>
    Private Function InitialStep() As String
        If Not PrivacyConsentGate.IsAccepted() Then Return "privacy"
        Return "example"
    End Function

    ''' <summary>
    ''' 示例任务话术只在 QuickActionCatalog 维护一份：这里曾经也有一份 Select Case，
    ''' 现在选区工具条（E4-01）、右键菜单（E4-02）、每端示例任务（E3-02）共用同一来源，
    ''' 避免三端各留一张需要人工同步的表。带 request 的那条是一键示例，点下去直接进主链。
    ''' </summary>
    Private Function BuildExamples() As JArray
        Dim arr As New JArray()
        For Each sample In ShareRibbon.Agent.QuickActionCatalog.SampleTasks(_getAppType())
            Dim item As New JObject()
            item("text") = sample.Text
            If sample.Runnable Then item("request") = sample.RunRequestText
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
