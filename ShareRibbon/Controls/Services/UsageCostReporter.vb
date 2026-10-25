Imports System.Globalization

''' <summary>
''' 用量的显示口径只有一个持有者：token 怎么变成给用户看的一行话、缺单价时说什么、超额度时说什么，
''' 都从这里出。终态卡片、流式消息脚注、日志读到的是同一句文案，不允许在 JS 里再拼一遍。
'''
''' 单价是用户自己填的（每百万 token），默认未配置。**未配置时只报 token，不报钱**——
''' 报一个猜出来的金额比不报更糟，因为它会被当成账单。
''' </summary>
Public Class UsageCostReporter

    Public Const UnpricedText As String = "未配置单价，只统计 token 用量"
    ''' <summary>provider 没回 usage 时说这句话，不回 0 tokens——0 会被读成"这次没花钱"。</summary>
    Public Const NoUsageText As String = "这次没有回传用量数据"
    ''' <summary>额度为 0 表示不限量，不是“一 token 就提醒”。</summary>
    Public Const BudgetUnlimited As Integer = 0

    ''' <summary>缺单价时说缺哪一项，而不是把没配好的单价按 0 算出一个免费的费用。</summary>
    Public Shared Function IncompletePricingWarning(filledSide As String) As String
        Return $"只填了{filledSide}单价，还缺另一侧：估算费用要输入和输出两个单价（每百万 token），这次先只显示 token 用量"
    End Function

    Public Shared Function CurrencyWithoutPricingWarning As String
        Return "填了货币符号就要同时填输入/输出单价（每百万 token），否则只显示 token 用量"
    End Function

    ''' <summary>面板输入框的口径：-1 与 0（不限量）都是内部哨兵值，显示成空而不是显示一个能改的假数字。</summary>
    Public Shared Function PriceInputText(value As Double) As String
        If value < 0 Then Return ""
        Return value.ToString(Globalization.CultureInfo.InvariantCulture)
    End Function

    Public Shared Function BudgetInputText(value As Integer) As String
        If value <= BudgetUnlimited Then Return ""
        Return value.ToString(Globalization.CultureInfo.InvariantCulture)
    End Function

    Public Shared Function IsPriced() As Boolean
        Return ChatSettings.PricePerMillionIn >= 0 AndAlso ChatSettings.PricePerMillionOut >= 0
    End Function

    ''' <summary>费用估算（单位由用户在设置里填的符号决定）；未配置单价时返回 -1，调用方不得当 0 用。</summary>
    Public Shared Function EstimateCost(promptTokens As Integer, completionTokens As Integer) As Double
        If Not IsPriced() Then Return -1
        Return (promptTokens / 1000000.0) * ChatSettings.PricePerMillionIn +
               (completionTokens / 1000000.0) * ChatSettings.PricePerMillionOut
    End Function

    ''' <summary>钱的那半句：没单价就没有这句话，而不是写个 0。</summary>
    Public Shared Function CostText(promptTokens As Integer, completionTokens As Integer) As String
        If Not IsPriced() Then Return ""
        Dim cost = EstimateCost(promptTokens, completionTokens)
        Return "估算费用 " & If(ChatSettings.PriceCurrency, "") &
               cost.ToString("0.####", CultureInfo.InvariantCulture)
    End Function

    ''' <summary>额度那句只在真的超了时出现；不限量或还没超都是空串。</summary>
    Public Shared Function BudgetText(sessionTokens As Integer) As String
        If ChatSettings.SessionTokenBudget <= BudgetUnlimited Then Return ""
        If sessionTokens < ChatSettings.SessionTokenBudget Then Return ""
        Return $"本轮会话已用 {sessionTokens} tokens，超过你设定的额度 {ChatSettings.SessionTokenBudget} tokens；可在设置里调整"
    End Function

    ''' <summary>给终态卡片和消息脚注用的一整句。</summary>
    Public Shared Function SummaryText(usage As TokenUsage, sessionTokens As Integer) As String
        If usage Is Nothing OrElse usage.TotalTokens <= 0 Then Return ""
        Dim parts As New List(Of String) From {
            $"{usage.TotalTokens} tokens"
        }
        Dim cost = CostText(usage.PromptTokens, usage.CompletionTokens)
        If cost <> "" Then
            parts.Add(cost)
        Else
            parts.Add(UnpricedText)
        End If
        Dim budget = BudgetText(sessionTokens)
        If budget <> "" Then parts.Add(budget)
        Return String.Join(" · ", parts)
    End Function
End Class
