' ShareRibbon\Controls\Services\PrivacyConsentGate.vb
' E3-05: 第一次把用户文档交给外部模型之前必须先有同意，而这份同意要能被查看和撤回。
' 数据边界三条文案只在这里持有：向导、面板设置、阻断卡片都从这一处读，前端不再抄一份。
' 判定与写入也收在这里，OnboardingService 只负责把状态推给面板，不再自己碰存储。

Imports ShareRibbon.Security

Public Class PrivacyConsentGate

    ''' <summary>面板 CTA 卡片的行为判别：卡片是共享的，按钮动作由载荷里的 action 决定。</summary>
    Public Const ActionAccept As String = "accept_privacy"

    ''' <summary>数据边界说明。改这里的文字等于改用户看到的承诺，因此全仓只许存在这一份。</summary>
    Public Shared ReadOnly BoundaryItems As String() = {
        "文档内容、选区和会话会发送给你所选的模型接口，用于生成结果。",
        "会话历史、变更日志和记忆只保存在本机 SQLite，不上传统计服务。",
        "API Key 只写入当前用户的加密存储，配置文件里只保留占位符。"
    }

    ''' <summary>所有出口共用同一条拒绝理由，避免出现第二种「为什么没发出去」的说法。</summary>
    Public Const RefusalReason As String = "privacy consent not accepted; request was not sent."

    Public Shared Function IsAccepted() As Boolean
        Return SecureConfigManager.PrivacySettings.PrivacyPolicyAccepted
    End Function

    ''' <summary>
    ''' 后台链路（补全、embedding、记忆沉淀、非流式工具请求）用的判定：这些调用点没有面板，
    ''' 拦下只能表现为「这次没有结果」，所以它们共用这一句日志，而不是各自编一条理由。
    ''' 面向用户的两条漏斗不走这里——它们要推卡片，见 BaseChatControl.IsBlockedByPrivacy。
    ''' </summary>
    Public Shared Function RefuseQuietly(callerName As String) As Boolean
        If IsAccepted() Then Return False
        AppLogger.Debug(callerName, RefusalReason)
        Return True
    End Function

    ''' <summary>同意与撤回都留一条记录：用户问「我什么时候点过同意」时，盘上要有答案。</summary>
    Public Shared Sub Apply(accepted As Boolean)
        SecureConfigManager.PrivacySettings.PrivacyPolicyAccepted = accepted
        AppLogger.Info("PrivacyConsentGate", $"privacy consent stored: accepted={If(accepted, "true", "false")}")
    End Sub

    ''' <summary>
    ''' 纯判定：把当前同意状态传进来。未同意才给卡片，且卡片必须说清「这一轮还没发出去」。
    ''' </summary>
    Public Shared Function Evaluate(accepted As Boolean, blockedAction As String) As ModelConfigPrompt
        If accepted Then Return Nothing

        Dim stopped = If(blockedAction = ModelConfigGate.ActionAgent,
                         "任务没有开始执行，文档不会被修改。",
                         "消息没有发送，文档不会被修改。")
        Return New ModelConfigPrompt With {
            .Reason = "privacy_not_accepted",
            .Action = ActionAccept,
            .Title = "先确认数据边界",
            .Message = String.Concat(BoundaryItems) & stopped & "确认后请重新发起这条请求；之后随时可以在面板设置里撤回。",
            .ActionLabel = "同意并继续",
            .BlockedAction = If(blockedAction, ModelConfigGate.ActionChat)
        }
    End Function
End Class
