' ShareRibbon\Agent\TaskPackageCatalog.Parsing.vb
' 上面那半个 TaskPackageCatalog 负责"从哪些目录读、读出来的东西放在哪"，
' 这一份只负责"一份 JSON 凭什么算合格"：字段白名单、枚举词表、id/文件名/宿主三者对齐、
' 断言事实与布尔/数值语义对齐、解释模板占位符只能引用真实存在的变更事实。
' 校验全在这里发生，门禁脚本（scripts/check-task-package-contract.ps1）才能把同一批规则
' 用内存里的坏样本反向跑一遍——两边任一边松手，另一边的探针就会变红。

Imports System.IO
Imports Newtonsoft.Json.Linq

Namespace Agent

    Partial Class TaskPackageCatalog

        Private Shared Function ParsePackage(node As JObject,
                                             sourceFile As String,
                                             errors As List(Of String)) As TaskPackage
            RejectUnknownKeys(node, AllowedTopLevelKeys, "包", errors)

            Dim pkg As New TaskPackage() With {.SourceFile = sourceFile}

            Dim versionToken = node("schemaVersion")
            If versionToken Is Nothing OrElse versionToken.Type <> JTokenType.Integer Then
                errors.Add("缺少整数字段 schemaVersion")
            ElseIf versionToken.Value(Of Integer)() <> AllowedSchemaVersion Then
                errors.Add($"schemaVersion={versionToken.Value(Of Integer)()}，加载器只认 {AllowedSchemaVersion}")
            End If

            pkg.Id = ReadRequiredText(node, "id", errors)
            pkg.Title = ReadRequiredText(node, "title", errors)
            pkg.Application = ReadRequiredText(node, "application", errors)
            pkg.Skill = ReadRequiredText(node, "skill", errors)
            pkg.Deliverable = ReadRequiredText(node, "deliverable", errors)
            pkg.EntryActions = ReadRequiredArray(node, "entryActions", errors)
            pkg.CapabilityIds = ReadOptionalArray(node, "capabilityIds")
            pkg.SuccessCriteria = ReadRequiredArray(node, "successCriteria", errors)
            pkg.GoldenCases = ReadOptionalArray(node, "goldenCases")

            Dim expectedPrefix = pkg.Id.Split("."c).FirstOrDefault()
            If Not String.IsNullOrWhiteSpace(pkg.Application) AndAlso
               Not AllowedApplications.Contains(pkg.Application) Then
                errors.Add($"application 必须是 {String.Join("/", AllowedApplications)} 之一，实际 {pkg.Application}")
            End If
            If Not String.IsNullOrWhiteSpace(pkg.Id) AndAlso
               Not String.IsNullOrWhiteSpace(pkg.Application) AndAlso
               Not IdPrefixMatchesApplication(expectedPrefix, pkg.Application) Then
                errors.Add($"包 id 前缀 {expectedPrefix} 与 application {pkg.Application} 不一致")
            End If
            If Not String.IsNullOrWhiteSpace(pkg.Id) AndAlso
               Not String.Equals(Path.GetFileNameWithoutExtension(sourceFile), pkg.Id, StringComparison.OrdinalIgnoreCase) Then
                errors.Add($"文件名 {sourceFile} 必须与包 id {pkg.Id} 同名")
            End If

            ParsePreviewPolicy(TryCast(node("previewPolicy"), JObject), pkg, errors)
            ParseVerifier(TryCast(node("verifier"), JObject), pkg, errors)
            ParseExplanation(TryCast(node("explanation"), JObject), pkg, errors)
            ParseRecovery(TryCast(node("recovery"), JObject), pkg, errors)
            ParseExistingImplementation(TryCast(node("existingImplementation"), JObject), pkg, errors)

            If pkg.EntryActions.Count > 0 Then
                Dim known = QuickActionCatalog.ActionsFor(pkg.Application).
                            Select(Function(action) action.Id).
                            ToList()
                For Each actionId In pkg.EntryActions
                    If Not known.Any(Function(existing) String.Equals(existing, actionId, StringComparison.OrdinalIgnoreCase)) Then
                        errors.Add($"入口动作 {actionId} 不在 QuickActionCatalog 的 {pkg.Application} 清单里")
                    End If
                Next
            End If

            Return pkg
        End Function

        Private Shared Function IdPrefixMatchesApplication(prefix As String, application As String) As Boolean
            Select Case application.Trim().ToLowerInvariant()
                Case "word"
                    Return String.Equals(prefix, "word", StringComparison.OrdinalIgnoreCase)
                Case "excel"
                    Return String.Equals(prefix, "excel", StringComparison.OrdinalIgnoreCase)
                Case "powerpoint"
                    Return String.Equals(prefix, "ppt", StringComparison.OrdinalIgnoreCase) OrElse
                           String.Equals(prefix, "powerpoint", StringComparison.OrdinalIgnoreCase)
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Sub ParsePreviewPolicy(node As JObject, pkg As TaskPackage, errors As List(Of String))
            If node Is Nothing Then
                errors.Add("缺少 previewPolicy：预览策略是本包必须承载的字段，不能靠人记")
                Return
            End If
            RejectUnknownKeys(node, AllowedPreviewKeys, "previewPolicy", errors)

            pkg.PreviewPolicy.Mode = ReadRequiredText(node, "mode", errors)
            Dim flag = node("requireBeforeWrite")
            If flag Is Nothing OrElse flag.Type <> JTokenType.Boolean Then
                errors.Add("previewPolicy.requireBeforeWrite 必须是布尔值")
            Else
                pkg.PreviewPolicy.RequireBeforeWrite = flag.Value(Of Boolean)()
            End If
            pkg.PreviewPolicy.Note = ReadRequiredText(node, "note", errors)

            If Not String.IsNullOrWhiteSpace(pkg.PreviewPolicy.Mode) AndAlso
               Not AllowedPreviewModes.Contains(pkg.PreviewPolicy.Mode) Then
                errors.Add($"previewPolicy.mode={pkg.PreviewPolicy.Mode} 不在词表 {String.Join("/", AllowedPreviewModes)} 中")
            End If
            ' 只读包却要求写前确认、或写包声称不预览却要写前确认，都是合同自相矛盾。
            If String.Equals(pkg.PreviewPolicy.Mode, "none", StringComparison.OrdinalIgnoreCase) AndAlso
               pkg.PreviewPolicy.RequireBeforeWrite Then
                errors.Add("previewPolicy.mode=none 时 requireBeforeWrite 必须为 false")
            End If
        End Sub

        Private Shared Sub ParseVerifier(node As JObject, pkg As TaskPackage, errors As List(Of String))
            If node Is Nothing Then
                errors.Add("缺少 verifier：没有验证断言的包无法回答「真的做成了吗」")
                Return
            End If
            RejectUnknownKeys(node, AllowedVerifierKeys, "verifier", errors)

            pkg.Verifier.OnFail = ReadRequiredText(node, "onFail", errors)
            If Not String.IsNullOrWhiteSpace(pkg.Verifier.OnFail) AndAlso
               Not AllowedOnFailStrategies.Contains(pkg.Verifier.OnFail) Then
                errors.Add($"verifier.onFail={pkg.Verifier.OnFail} 不在词表 {String.Join("/", AllowedOnFailStrategies)} 中")
            End If

            Dim items = TryCast(node("assertions"), JArray)
            If items Is Nothing OrElse items.Count = 0 Then
                errors.Add("verifier.assertions 至少要有 1 条断言")
                Return
            End If

            For Each raw In items.OfType(Of JObject)()
                RejectUnknownKeys(raw, AllowedAssertionKeys, "断言", errors)
                Dim assertion As New TaskPackageAssertion With {
                    .Fact = ReadRequiredText(raw, "fact", errors),
                    .Op = ReadRequiredText(raw, "op", errors),
                    .Why = ReadRequiredText(raw, "why", errors)
                }

                If Not String.IsNullOrWhiteSpace(assertion.Fact) AndAlso Not VerifiableFacts.Contains(assertion.Fact) Then
                    errors.Add($"断言事实 {assertion.Fact} 不是 AgentRunChangeLog 上的可核对属性")
                    Continue For
                End If
                If Not String.IsNullOrWhiteSpace(assertion.Op) AndAlso Not AllowedOps.Contains(assertion.Op) Then
                    errors.Add($"断言 op={assertion.Op} 不在词表 {String.Join("/", AllowedOps)} 中")
                End If

                Dim isBooleanFact = String.Equals(assertion.Fact, "hasChanges", StringComparison.OrdinalIgnoreCase) OrElse
                                    String.Equals(assertion.Fact, "nativeUndoable", StringComparison.OrdinalIgnoreCase) OrElse
                                    String.Equals(assertion.Fact, "fullyUndoable", StringComparison.OrdinalIgnoreCase)
                If isBooleanFact Then
                    If assertion.Op <> "isTrue" AndAlso assertion.Op <> "isFalse" Then
                        errors.Add($"布尔事实 {assertion.Fact} 只能用 isTrue/isFalse，实际 {assertion.Op}")
                    End If
                ElseIf assertion.Op = "isTrue" OrElse assertion.Op = "isFalse" Then
                    errors.Add($"数值事实 {assertion.Fact} 不能用 {assertion.Op}")
                Else
                    Dim valueToken = raw("value")
                    If valueToken Is Nothing OrElse valueToken.Type <> JTokenType.Integer Then
                        errors.Add($"断言 {assertion.Fact} 的 op={assertion.Op} 需要整数 value")
                    Else
                        assertion.Value = valueToken.Value(Of Integer)()
                    End If
                End If

                pkg.Verifier.Assertions.Add(assertion)
            Next
        End Sub

        Private Shared Sub ParseExplanation(node As JObject, pkg As TaskPackage, errors As List(Of String))
            If node Is Nothing Then
                errors.Add("缺少 explanation：解释模板是本包必须承载的字段")
                Return
            End If
            RejectUnknownKeys(node, AllowedExplanationKeys, "explanation", errors)
            pkg.Explanation.WhatChanged = ReadRequiredText(node, "whatChanged", errors)
            pkg.Explanation.WhyHint = ReadRequiredText(node, "whyHint", errors)
            pkg.Explanation.UndoHint = ReadRequiredText(node, "undoHint", errors)

            For Each field In New List(Of String) From {
                    pkg.Explanation.WhatChanged, pkg.Explanation.WhyHint, pkg.Explanation.UndoHint}
                For Each token In PlaceholderTokens(field)
                    If Not VerifiableFacts.Contains(token) Then
                        errors.Add($"解释模板里的占位符 {{{token}}} 不是可核对的变更事实")
                    End If
                Next
            Next
        End Sub

        Private Shared Sub ParseRecovery(node As JObject, pkg As TaskPackage, errors As List(Of String))
            If node Is Nothing Then
                errors.Add("缺少 recovery：失败恢复口径是本包必须承载的字段")
                Return
            End If
            RejectUnknownKeys(node, AllowedRecoveryKeys, "recovery", errors)
            pkg.Recovery.Strategy = ReadRequiredText(node, "strategy", errors)
            pkg.Recovery.Notes = ReadRequiredText(node, "notes", errors)
            If Not String.IsNullOrWhiteSpace(pkg.Recovery.Strategy) AndAlso
               Not AllowedRecoveryStrategies.Contains(pkg.Recovery.Strategy) Then
                errors.Add($"recovery.strategy={pkg.Recovery.Strategy} 不在词表 {String.Join("/", AllowedRecoveryStrategies)} 中")
            End If
        End Sub

        Private Shared Sub ParseExistingImplementation(node As JObject, pkg As TaskPackage, errors As List(Of String))
            If node Is Nothing Then
                errors.Add("缺少 existingImplementation：必须先说清仓库里已经有什么，再声明补什么")
                Return
            End If
            RejectUnknownKeys(node, AllowedImplementationKeys, "existingImplementation", errors)
            pkg.ExistingImplementation.Reuse = ReadRequiredArray(node, "reuse", errors)
            pkg.ExistingImplementation.Gap = ReadRequiredText(node, "gap", errors)
        End Sub

        Private Shared Function PlaceholderTokens(text As String) As List(Of String)
            Dim result As New List(Of String)()
            If String.IsNullOrEmpty(text) Then Return result
            For Each match In System.Text.RegularExpressions.Regex.Matches(text, "\{([A-Za-z][A-Za-z0-9_]*)\}")
                result.Add(match.Groups(1).Value)
            Next
            Return result
        End Function

        Private Shared Sub RejectUnknownKeys(node As JObject,
                                            allowed As IReadOnlyCollection(Of String),
                                            where As String,
                                            errors As List(Of String))
            For Each prop In node.Properties()
                If Not allowed.Contains(prop.Name) Then
                    errors.Add($"{where} 出现未知字段 {prop.Name}（任务包 schema 只认 {String.Join("/", allowed)}）")
                End If
            Next
        End Sub

        Private Shared Function ReadRequiredText(node As JObject, name As String, errors As List(Of String)) As String
            Dim token = node(name)
            If token Is Nothing OrElse token.Type <> JTokenType.String OrElse String.IsNullOrWhiteSpace(token.Value(Of String)()) Then
                errors.Add($"{name} 必须是非空字符串")
                Return ""
            End If
            Return token.Value(Of String)().Trim()
        End Function

        Private Shared Function ReadRequiredArray(node As JObject, name As String, errors As List(Of String)) As List(Of String)
            Dim items = ReadOptionalArray(node, name)
            If items.Count = 0 Then errors.Add($"{name} 至少要有 1 项")
            Return items
        End Function

        Private Shared Function ReadOptionalArray(node As JObject, name As String) As List(Of String)
            Dim result As New List(Of String)()
            Dim items = TryCast(node(name), JArray)
            If items Is Nothing Then Return result
            For Each item In items.OfType(Of JValue)().Where(Function(value) value.Type = JTokenType.String)
                Dim text = item.Value(Of String)()
                If Not String.IsNullOrWhiteSpace(text) Then result.Add(text.Trim())
            Next
            Return result
        End Function
    End Class

End Namespace
