' ShareRibbon\Agent\TaskPackageOutcomeVerifier.vb
' E6-05：任务包 verifier.assertions 在 run 收尾时的真核对。
' 事实一律从 AgentRunChangeLog.TryFactValue 读——台账是变更事实的唯一持有者，本类不自己数记录、也不新增事实。

Imports System.Collections.Generic

Namespace Agent

    ''' <summary>
    ''' 把包声明的成功标准变成一条条能判真假的断言。判定失败即"未达成"，调用方据此不能宣称成功。
    ''' </summary>
    Public NotInheritable Class TaskPackageOutcomeVerifier

        Private Sub New()
        End Sub

        ''' <summary>
        ''' 未达成的断言列表；全部达成时返回空表。
        ''' 读不出的事实、不认识的比较符、类型对不上都算未达成：宁可把一次成功说成未完成，
        ''' 也不能让断言在拼错时静默通过——那等于没有验证器。
        ''' </summary>
        Public Shared Function UnmetReasons(pkg As TaskPackage, log As AgentRunChangeLog) As List(Of String)
            Dim reasons As New List(Of String)()
            Dim assertions = If(pkg?.Verifier?.Assertions, Nothing)
            If assertions Is Nothing OrElse assertions.Count = 0 Then Return reasons
            For Each assertion In assertions
                Dim reason = CheckOne(assertion, log)
                If Not String.IsNullOrWhiteSpace(reason) Then reasons.Add(reason)
            Next
            Return reasons
        End Function

        ''' <summary>断言里声明的失败处置；没有 verifier 时返回空串，调用方按"没有合同"处理。</summary>
        Public Shared Function OnFail(pkg As TaskPackage) As String
            Return If(pkg?.Verifier?.OnFail, "").Trim()
        End Function

        Private Shared Function CheckOne(assertion As TaskPackageAssertion, log As AgentRunChangeLog) As String
            Dim fact = If(assertion?.Fact, "").Trim()
            If fact.Length = 0 Then Return "断言没有写事实名"
            If log Is Nothing Then Return FailText(assertion, fact, "本轮没有变更台账，无从核对")

            Dim actual As Object = Nothing
            If Not log.TryFactValue(fact, actual) Then Return FailText(assertion, fact, "台账里没有这个事实")

            Dim op = If(assertion.Op, "").Trim()
            Select Case op
                Case "eq", "gte", "lte"
                    Dim asNumber As Long
                    If Not TryToLong(actual, asNumber) Then
                        Return FailText(assertion, fact, $"{Describe(actual)} 不是数值，不能用 {op} 比较")
                    End If
                    Dim passed As Boolean
                    Select Case op
                        Case "eq" : passed = (asNumber = assertion.Value)
                        Case "gte" : passed = (asNumber >= assertion.Value)
                        Case Else : passed = (asNumber <= assertion.Value)
                    End Select
                    If passed Then Return ""
                    Return FailText(assertion, fact, $"实际 {asNumber}")
                Case "isTrue", "isFalse"
                    If Not (IsBoolean(actual)) Then
                        Return FailText(assertion, fact, $"{Describe(actual)} 不是布尔值，不能用 {op} 判断")
                    End If
                    Dim expected = String.Equals(op, "isTrue", StringComparison.Ordinal)
                    If DirectCast(actual, Boolean) = expected Then Return ""
                    Return FailText(assertion, fact, $"实际 {If(DirectCast(actual, Boolean), "true", "false")}")
                Case Else
                    Return FailText(assertion, fact, $"比较符 {If(op.Length = 0, "(空)", op)} 不在词表里")
            End Select
        End Function

        ''' <summary>失败文案带上声明值与 why，用户才知道"包本来要求什么"，而不是一句空泛的未完成。</summary>
        Private Shared Function FailText(assertion As TaskPackageAssertion, fact As String, detail As String) As String
            Dim op = If(assertion?.Op, "").Trim()
            Dim expected = op
            If op = "eq" OrElse op = "gte" OrElse op = "lte" Then expected &= " " & assertion.Value
            Dim head = $"{fact} {expected}：{detail}"
            Dim why = If(assertion?.Why, "").Trim()
            If why.Length = 0 Then Return head
            Return head & "（" & why & "）"
        End Function

        Private Shared Function TryToLong(value As Object, ByRef result As Long) As Boolean
            result = 0
            If TypeOf value Is Integer OrElse TypeOf value Is Long Then
                result = Convert.ToInt64(value)
                Return True
            End If
            Return False
        End Function

        Private Shared Function IsBoolean(value As Object) As Boolean
            Return TypeOf value Is Boolean
        End Function

        Private Shared Function Describe(value As Object) As String
            If value Is Nothing Then Return "空值"
            Return value.GetType().Name
        End Function

    End Class

End Namespace
