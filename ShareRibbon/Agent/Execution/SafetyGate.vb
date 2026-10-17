Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon.Agent.OfficeOperations

Namespace Agent.Execution

    Public Enum SafetyAction
        Allow
        RequireApproval
        Deny
    End Enum

    Public Class SafetyDecision
        Public Property Action As SafetyAction = SafetyAction.Allow
        Public Property Reason As String = ""
        Public Property UserMessage As String = ""
        Public Property ErrorCode As String = ""
        Public Property RiskLevel As String = "safe"

        Public Shared Function Allow(Optional riskLevel As String = "safe") As SafetyDecision
            Return New SafetyDecision With {
                .Action = SafetyAction.Allow,
                .RiskLevel = If(riskLevel, "safe")
            }
        End Function

        Public Shared Function RequireApproval(reason As String,
                                               Optional userMessage As String = Nothing,
                                               Optional riskLevel As String = "risky") As SafetyDecision
            Return New SafetyDecision With {
                .Action = SafetyAction.RequireApproval,
                .Reason = If(reason, ""),
                .UserMessage = If(userMessage, reason),
                .ErrorCode = ExceptionClassifier.CodeSafetyNeedsApproval,
                .RiskLevel = If(riskLevel, "risky")
            }
        End Function

        Public Shared Function Deny(reason As String,
                                    errorCode As String,
                                    Optional userMessage As String = Nothing,
                                    Optional riskLevel As String = "risky") As SafetyDecision
            Return New SafetyDecision With {
                .Action = SafetyAction.Deny,
                .Reason = If(reason, ""),
                .UserMessage = If(userMessage, reason),
                .ErrorCode = If(errorCode, ExceptionClassifier.CodeSafetyBlocked),
                .RiskLevel = If(riskLevel, "risky")
            }
        End Function
    End Class

    Public Class SafetyGate
        Public Property VbaEnabled As Boolean = False
        Public Property RequireApprovalForRisky As Boolean = True
        Public Property RequireApprovalForDelete As Boolean = True

        Public Function Evaluate(tool As ToolDescriptor, params As JObject) As SafetyDecision
            If tool Is Nothing Then
                Return SafetyDecision.Deny("工具不存在，无法执行安全裁决",
                                           ExceptionClassifier.CodeNotFound,
                                           "工具不存在，无法执行")
            End If

            Dim toolId = If(tool.Id, "")
            Dim risk = If(String.IsNullOrWhiteSpace(tool.RiskLevel), "risky", tool.RiskLevel.Trim().ToLowerInvariant())

            If tool.IsVbaFallback OrElse String.Equals(toolId, "ExecuteVBA", StringComparison.OrdinalIgnoreCase) Then
                Return EvaluateVba(toolId, params, risk)
            End If

            If String.Equals(toolId, "OfficeObjectOperation", StringComparison.OrdinalIgnoreCase) Then
                Return EvaluateOfficeOperation(params)
            End If

            If RequireApprovalForDelete AndAlso IsDestructiveTool(toolId, params) Then
                Return SafetyDecision.RequireApproval($"工具 {toolId} 可能删除或清空内容",
                                                      $"工具 {toolId} 需要确认后才能执行",
                                                      "risky")
            End If

            If RequireApprovalForRisky AndAlso String.Equals(risk, "risky", StringComparison.OrdinalIgnoreCase) Then
                Return SafetyDecision.RequireApproval($"高风险工具 {toolId} 需要用户确认",
                                                      $"高风险工具 {toolId} 需要确认后才能执行",
                                                      risk)
            End If

            Return SafetyDecision.Allow(risk)
        End Function

        Private Function EvaluateOfficeOperation(params As JObject) As SafetyDecision
            Dim batchToken = params?("batch")
            If batchToken Is Nothing OrElse batchToken.Type <> JTokenType.Object Then
                Return SafetyDecision.Deny("OfficeObjectOperation 缺少合法 batch",
                                           ExceptionClassifier.CodeOperationSchemaInvalid,
                                           "声明式 Office 操作格式无效")
            End If

            Dim batch As OfficeOperationBatch = Nothing
            Try
                batch = batchToken.ToObject(Of OfficeOperationBatch)()
            Catch ex As Exception
                Return SafetyDecision.Deny("OfficeObjectOperation batch 反序列化失败",
                                           ExceptionClassifier.CodeOperationSchemaInvalid,
                                           "声明式 Office 操作格式无效")
            End Try

            Dim validation = OfficeOperationValidation.ValidateBatch(batch)
            If Not validation.IsValid Then
                Return SafetyDecision.Deny(validation.ToErrorMessage(),
                                           ExceptionClassifier.CodeOperationSchemaInvalid,
                                           "声明式 Office 操作未通过合同校验")
            End If

            Dim requiresApproval As Boolean = False
            Dim highestRisk As String = "safe"
            For Each operation In batch.Operations
                Dim action = If(operation.Action, "").Trim().ToLowerInvariant()
                Dim memberId = If(operation.MemberId, "").Trim().ToLowerInvariant()

                If ContainsMemberToken(memberId, "quit") OrElse
                   ContainsMemberToken(memberId, "vbproject") OrElse
                   ContainsMemberToken(memberId, "shell") OrElse
                   ContainsMemberToken(memberId, "run") OrElse
                   ContainsMemberToken(memberId, "executemso") Then
                    Return SafetyDecision.Deny($"成员 {operation.MemberId} 禁止通过声明式操作执行",
                                               ExceptionClassifier.CodeSafetyBlocked,
                                               "该 Office API 成员不允许执行")
                End If

                If action = "delete" OrElse
                   ContainsMemberToken(memberId, "delete") OrElse
                   ContainsMemberToken(memberId, "remove") OrElse
                   ContainsMemberToken(memberId, "clear") OrElse
                   ContainsMemberToken(memberId, "close") OrElse
                   ContainsMemberToken(memberId, "saveas") OrElse
                   ContainsMemberToken(memberId, "savecopyas") OrElse
                   ContainsMemberToken(memberId, "save") OrElse
                   ContainsMemberToken(memberId, "export") OrElse
                   ContainsMemberToken(memberId, "printout") OrElse
                   ContainsMemberToken(memberId, "sendmail") OrElse
                   ContainsMemberToken(memberId, "checkin") OrElse
                   ContainsMemberToken(memberId, "checkout") OrElse
                   ContainsMemberToken(memberId, "reload") OrElse
                   ContainsMemberToken(memberId, "protect") Then
                    requiresApproval = True
                    highestRisk = "risky"
                ElseIf action <> "get" Then
                    If highestRisk <> "risky" Then highestRisk = "medium"
                End If
            Next

            If requiresApproval Then
                Return SafetyDecision.RequireApproval("声明式 Office 操作包含删除、关闭或覆盖类成员",
                                                      "该 Office 操作可能删除、关闭或覆盖内容，需要确认后执行",
                                                      highestRisk)
            End If
            Return SafetyDecision.Allow(highestRisk)
        End Function

        Private Shared Function ContainsMemberToken(memberId As String, memberName As String) As Boolean
            If String.IsNullOrWhiteSpace(memberId) OrElse String.IsNullOrWhiteSpace(memberName) Then Return False
            Return memberId.IndexOf("." & memberName & "(", StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        Private Function EvaluateVba(toolId As String, params As JObject, risk As String) As SafetyDecision
            If Not VbaEnabled Then
                Return SafetyDecision.Deny("VBA 工具默认关闭",
                                           ExceptionClassifier.CodeVbaDisabled,
                                           "VBA 执行默认关闭，未进入宿主执行器",
                                           risk)
            End If

            Dim code = If(params?("code")?.ToString(), "")
            Dim subResult = SafetyChecker.Check(code)
            If subResult IsNot Nothing AndAlso Not subResult.IsSafe Then
                Return SafetyDecision.Deny(subResult.Reason,
                                           ExceptionClassifier.CodeSafetyBlocked,
                                           subResult.ToUserMessage(),
                                           risk)
            End If

            If subResult IsNot Nothing AndAlso subResult.NeedsConfirm Then
                Return SafetyDecision.RequireApproval(subResult.Reason,
                                                      subResult.ToUserMessage(),
                                                      risk)
            End If

            Return SafetyDecision.RequireApproval($"VBA 工具 {toolId} 需要用户确认",
                                                  $"VBA 工具 {toolId} 需要确认后才能执行",
                                                  risk)
        End Function

        Private Function IsDestructiveTool(toolId As String, params As JObject) As Boolean
            Dim id = If(toolId, "").ToLowerInvariant()
            If id.Contains("delete") OrElse id.Contains("clear") OrElse id.Contains("remove") Then Return True

            If String.Equals(id, "replacetext", StringComparison.OrdinalIgnoreCase) Then
                Dim rangeName = If(params?("range")?.ToString(), "all")
                Return IsWholeDocumentRange(rangeName)
            End If

            Dim scope = If(params?("scope")?.ToString(), "")
            Dim range = If(params?("range")?.ToString(), "")
            Return IsWholeDocumentRange(scope) OrElse IsWholeDocumentRange(range)
        End Function

        Private Function IsWholeDocumentRange(value As String) As Boolean
            Dim normalized = If(value, "").Trim().ToLowerInvariant()
            Return normalized = "all" OrElse
                   normalized = "document" OrElse
                   normalized = "workbook" OrElse
                   normalized = "presentation" OrElse
                   normalized = "全文"
        End Function
    End Class

End Namespace
