' ShareRibbon\Controls\Services\ChatContextBuilder.vb
' 分层上下文组装：[0]～[6]

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text

''' <summary>
''' Chat 上下文构建器：按 roadmap 2.5 分层组装消息
''' </summary>
Public Class ChatContextBuilder
    <ThreadStatic>
    Private Shared _lastTrace As ChatContextTrace

    Public Shared ReadOnly Property LastTrace As ChatContextTrace
        Get
            Return _lastTrace
        End Get
    End Property

    ''' <summary>
    ''' 构建分层消息列表，各层之间以结构化标题隔开，便于 AI 理解上下文来源。
    ''' </summary>
    ''' <param name="scenario">excel/word/ppt/common</param>
    ''' <param name="appType">当前宿主类型</param>
    ''' <param name="currentQuery">用户当前输入（用于 RAG）</param>
    ''' <param name="sessionMessages">当前会话滚动窗口 (user/assistant)</param>
    ''' <param name="latestUserMessage">本条 user 消息</param>
    ''' <param name="baseSystemPrompt">已有 system 提示词（来自 PromptManager 等）</param>
    ''' <param name="variableValues">变量替换字典，如 {{选中内容}}</param>
    ''' <param name="enableMemory">是否启用 Memory（RAG、用户画像、会话摘要）</param>
    ''' <param name="ragCountOut">输出：本次检索到的记忆条数（供 UI 显示，避免调用方再次查询）</param>
    ''' <returns>按 [0]～[6] 顺序的消息列表</returns>
    Public Shared Function BuildMessages(
        scenario As String,
        appType As String,
        currentQuery As String,
        sessionMessages As List(Of HistoryMessage),
        latestUserMessage As String,
        baseSystemPrompt As String,
        variableValues As Dictionary(Of String, String),
        enableMemory As Boolean,
        Optional ByRef ragCountOut As Integer = 0) As List(Of HistoryMessage)

        Dim result As New List(Of HistoryMessage)()
        ragCountOut = 0
        Dim scenarioNorm = If(String.IsNullOrEmpty(scenario), "common", scenario.ToLowerInvariant())
        Dim appNorm = If(String.IsNullOrEmpty(appType), "Excel", appType)
        Dim vars = If(variableValues, New Dictionary(Of String, String)())
        Dim trace As New ChatContextTrace With {
            .Query = currentQuery,
            .AppType = appNorm
        }
        _lastTrace = trace

        ' 所有 system 层收集到 sysParts，最终合并为单一 system 消息，节之间用 --- 分隔
        Dim sysParts As New List(Of String)()

        ' [0] 角色与基础指令
        If Not String.IsNullOrWhiteSpace(baseSystemPrompt) Then
            sysParts.Add("### 角色与基础指令" & vbCrLf & baseSystemPrompt.Trim())
        End If

        ' [1] 场景能力（数据库场景提示词）
        Dim systemPromptFromDb = PromptTemplateRepository.GetSystemPrompt(scenarioNorm)
        If Not String.IsNullOrWhiteSpace(systemPromptFromDb) Then
            sysParts.Add("### 场景能力" & vbCrLf & PromptTemplateRepository.ReplaceVariables(systemPromptFromDb.Trim(), vars))
        End If

        ' [1b] 可用技能（Skills 渐进式披露）
        Dim skillsCatalog = SkillsService.GetSkillsCatalog()
        If skillsCatalog IsNot Nothing AndAlso skillsCatalog.Count > 0 Then
            Dim skillParts As New List(Of String)()
            skillParts.Add("### 可用技能")

            Dim catalogMessage = SkillsService.BuildSkillsCatalogMessage(skillsCatalog)
            If Not String.IsNullOrWhiteSpace(catalogMessage) Then
                skillParts.Add(catalogMessage)
            End If

            Dim indexedSkills = SkillsIndexService.SelectSkillDefinitions(currentQuery, Nothing, appNorm, 2)
            If indexedSkills IsNot Nothing AndAlso indexedSkills.Count > 0 Then
                For Each indexedSkill In indexedSkills
                    Dim detailMessage = SkillsService.BuildSkillDetailMessage(indexedSkill)
                    If Not String.IsNullOrWhiteSpace(detailMessage) Then
                        skillParts.Add("#### 推荐技能（基于索引召回）")
                        skillParts.Add(detailMessage)
                    End If
                    AppendSkillScriptInfo(skillParts, indexedSkill)
                    skillParts.Add($"> 当前推荐: {indexedSkill.Name}")
                    trace.Skills.Add(New ChatContextSkillTrace With {
                        .Name = indexedSkill.Name,
                        .Source = "index",
                        .Reason = "基于当前查询召回"
                    })
                    Debug.WriteLine($"[ChatContextBuilder] 索引召回Skill: {indexedSkill.Name}")
                Next
            Else
                Dim matchedSkills = SkillsService.MatchSkills(currentQuery, 5)
                If matchedSkills.Count > 0 Then
                    Dim topSkill = matchedSkills.First()
                    If topSkill.MatchScore >= 10 Then
                        Dim topSkillDetail = SkillsDirectoryService.LoadSkillDetail(topSkill.Skill)
                        If topSkillDetail Is Nothing Then topSkillDetail = topSkill.Skill

                        Dim detailMessage = SkillsService.BuildSkillDetailMessage(topSkillDetail)
                        If Not String.IsNullOrWhiteSpace(detailMessage) Then
                            skillParts.Add("#### 推荐技能（基于当前查询）")
                            skillParts.Add(detailMessage)
                        End If

                        ' 如果 Skill 有脚本，添加脚本信息
                        AppendSkillScriptInfo(skillParts, topSkillDetail)

                        Dim metaHints As New List(Of String)()
                        metaHints.Add($"当前推荐: {topSkillDetail.Name}")
                        If topSkillDetail.Tags IsNot Nothing AndAlso topSkillDetail.Tags.Count > 0 Then
                            metaHints.Add($"标签: {String.Join(", ", topSkillDetail.Tags)}")
                        End If
                        If Not String.IsNullOrWhiteSpace(topSkillDetail.Compatibility) Then
                            metaHints.Add($"兼容性: {topSkillDetail.Compatibility}")
                        End If
                        If topSkill.MatchedKeywords.Count > 0 Then
                            metaHints.Add($"匹配关键词: {String.Join(", ", topSkill.MatchedKeywords)}")
                        End If
                        skillParts.Add("> " & String.Join(" | ", metaHints))
                        trace.Skills.Add(New ChatContextSkillTrace With {
                            .Name = topSkillDetail.Name,
                            .Source = "keyword",
                            .Reason = If(topSkill.MatchedKeywords.Count > 0, String.Join(", ", topSkill.MatchedKeywords), "基于当前查询匹配")
                        })

                        Debug.WriteLine($"[ChatContextBuilder] 匹配到Skill: {topSkillDetail.Name}, 分数: {topSkill.MatchScore:F1}, 关键词: {String.Join(", ", topSkill.MatchedKeywords)}")
                    End If
                Else
                    Debug.WriteLine($"[ChatContextBuilder] 未匹配到Skills，提供 {skillsCatalog.Count} 个Skill目录")
                End If
            End If

            sysParts.Add(String.Join(vbCrLf & vbCrLf, skillParts))
        End If

        ' [3][4] 用户上下文：画像 + RAG 记忆 + 近期会话摘要（仅一次检索）
        If enableMemory Then
            Debug.WriteLine("[ChatContextBuilder] 启用记忆，开始检索...")
            Dim memParts As New List(Of String)()
            memParts.Add("### 用户上下文")

            Dim userProfile = MemoryService.GetUserProfile()
            If Not String.IsNullOrWhiteSpace(userProfile) Then
                Debug.WriteLine("[ChatContextBuilder] 找到用户画像")
                trace.UserProfileInjected = True
                memParts.Add("#### 用户画像" & vbCrLf & userProfile.Trim())
            End If

            Dim structuredMemories = MemoryService.GetRelevantStructuredMemories(currentQuery, Nothing, appNorm)
            If structuredMemories IsNot Nothing AndAlso structuredMemories.Count > 0 Then
                structuredMemories = structuredMemories.
                    Where(Function(m) IsRelevantToQuery(currentQuery, If(m.Content, "") & " " & If(m.Summary, ""))).
                    ToList()
            End If
            If structuredMemories IsNot Nothing AndAlso structuredMemories.Count > 0 Then
                ragCountOut = structuredMemories.Count
                Debug.WriteLine($"[ChatContextBuilder] 找到 {structuredMemories.Count} 条结构化记忆")
                Dim memLines As New List(Of String)()
                memLines.Add("#### 相关记忆")
                For Each m In structuredMemories
                    Dim label = If(String.IsNullOrWhiteSpace(m.MemoryType), "memory", m.MemoryType)
                    memLines.Add($"- [{label}] {m.Content}")
                    trace.Memories.Add(New ChatContextMemoryTrace With {
                        .Id = m.MemoryId,
                        .Source = "structured",
                        .MemoryType = label,
                        .Content = m.Content,
                        .Score = m.Score,
                        .Importance = m.Importance
                    })
                Next
                memParts.Add(String.Join(vbCrLf, memLines))
            Else
                Dim memories = MemoryService.GetRelevantMemories(currentQuery, Nothing, Nothing, Nothing, appNorm)
                If memories IsNot Nothing AndAlso memories.Count > 0 Then
                    memories = memories.
                        Where(Function(m) IsRelevantToQuery(currentQuery, If(m.Content, "") & " " & If(m.Tags, ""))).
                        ToList()
                End If
                If memories IsNot Nothing AndAlso memories.Count > 0 Then
                    ragCountOut = memories.Count
                    Debug.WriteLine($"[ChatContextBuilder] 找到 {memories.Count} 条原子记忆")
                    Dim memLines As New List(Of String)()
                    memLines.Add("#### 相关记忆")
                    For Each m In memories
                        memLines.Add("- " & m.Content)
                        trace.Memories.Add(New ChatContextMemoryTrace With {
                            .Id = m.Id.ToString(),
                            .Source = "atomic",
                            .MemoryType = If(String.IsNullOrWhiteSpace(m.MemoryType), "long_term", m.MemoryType),
                            .Content = m.Content,
                            .Score = m.SimilarityScore,
                            .Importance = m.Importance
                        })
                    Next
                    memParts.Add(String.Join(vbCrLf, memLines))
                Else
                    Dim q = If(currentQuery, "")
                    Debug.WriteLine($"[ChatContextBuilder] 没有找到相关记忆，查询: {q.Substring(0, Math.Min(100, q.Length))}...")
                End If
            End If

            Dim summaries = MemoryService.GetRecentSessionSummaries(Nothing, appNorm)
            If summaries IsNot Nothing AndAlso summaries.Count > 0 Then
                summaries = summaries.
                    Where(Function(s) IsRelevantToQuery(currentQuery, If(s.Title, "") & " " & If(s.Snippet, ""))).
                    Take(2).
                    ToList()
            End If
            If summaries IsNot Nothing AndAlso summaries.Count > 0 Then
                Debug.WriteLine($"[ChatContextBuilder] 找到 {summaries.Count} 条相关近期会话")
                Dim sumLines As New List(Of String)()
                sumLines.Add("#### 相关近期会话")
                For Each s In summaries
                    sumLines.Add($"- {s.Title}: {s.Snippet}")
                    trace.RecentSessions.Add(New ChatContextSessionTrace With {
                        .SessionId = s.SessionId,
                        .Title = s.Title,
                        .Snippet = s.Snippet
                    })
                Next
                memParts.Add(String.Join(vbCrLf, sumLines))
            End If

            ' 只有有实质内容（>1 表示除标题外至少有一项）时才注入
            If memParts.Count > 1 Then
                Debug.WriteLine($"[ChatContextBuilder] 组装记忆块，共 {memParts.Count - 1} 项")
                sysParts.Add(String.Join(vbCrLf & vbCrLf, memParts))
            Else
                Debug.WriteLine("[ChatContextBuilder] 没有记忆内容可注入")
            End If
        Else
            Debug.WriteLine("[ChatContextBuilder] 记忆被禁用")
        End If

        ' 将所有 system 层合并为单一消息，节之间用 --- 分隔
        If sysParts.Count > 0 Then
            Dim sep = vbCrLf & vbCrLf & "---" & vbCrLf & vbCrLf
            result.Insert(0, New HistoryMessage With {
                .role = "system",
                .content = String.Join(sep, sysParts)
            })
        End If

        ' [5] 当前会话滚动窗口（只含 user/assistant）
        If sessionMessages IsNot Nothing Then
            Dim addedCount = 0
            For Each msg In sessionMessages
                If msg.role <> "system" AndAlso Not String.IsNullOrEmpty(msg.content) Then
                    result.Add(New HistoryMessage With {.role = msg.role, .content = msg.content})
                    addedCount += 1
                End If
            Next
            Debug.WriteLine($"[ChatContextBuilder] 会话窗口添加 {addedCount} 条消息")
        End If

        ' [6] 本条 user 消息
        If Not String.IsNullOrWhiteSpace(latestUserMessage) Then
            result.Add(New HistoryMessage With {.role = "user", .content = latestUserMessage})
        End If

        Debug.WriteLine($"[ChatContextBuilder] 构建完成，消息数: {result.Count}，RAG命中: {ragCountOut}")
        Return result
    End Function

    Private Shared Function IsRelevantToQuery(query As String, text As String) As Boolean
        If String.IsNullOrWhiteSpace(query) OrElse String.IsNullOrWhiteSpace(text) Then Return False

        Dim queryNorm = NormalizeForMatch(query)
        Dim textNorm = NormalizeForMatch(text)
        If queryNorm.Length < 2 OrElse textNorm.Length < 2 Then Return False
        If textNorm.Contains(queryNorm) Then Return True

        Dim tokens = BuildQueryTokens(queryNorm)
        If tokens.Count = 0 Then Return False

        Dim hits = 0
        For Each token In tokens
            If textNorm.Contains(token) Then
                hits += 1
                If hits >= 1 Then Return True
            End If
        Next

        Return False
    End Function

    Private Shared Function NormalizeForMatch(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""
        Dim sb As New StringBuilder()
        For Each ch In value.Trim().ToLowerInvariant()
            If Char.IsLetterOrDigit(ch) OrElse IsCjk(ch) Then
                sb.Append(ch)
            Else
                sb.Append(" "c)
            End If
        Next
        Return String.Join(" ", sb.ToString().Split({" "c}, StringSplitOptions.RemoveEmptyEntries))
    End Function

    Private Shared Function BuildQueryTokens(queryNorm As String) As List(Of String)
        Dim tokens As New List(Of String)()
        For Each part In queryNorm.Split({" "c}, StringSplitOptions.RemoveEmptyEntries)
            If part.Length >= 2 AndAlso Not IsWeakQueryToken(part) Then AddToken(tokens, part)

            If ContainsCjk(part) AndAlso part.Length >= 3 Then
                For i = 0 To part.Length - 2
                    Dim gram = part.Substring(i, 2)
                    If Not IsWeakQueryToken(gram) Then AddToken(tokens, gram)
                Next
            End If
        Next
        Return tokens.Take(16).ToList()
    End Function

    Private Shared Sub AddToken(tokens As List(Of String), token As String)
        If String.IsNullOrWhiteSpace(token) Then Return
        If tokens.Any(Function(t) String.Equals(t, token, StringComparison.OrdinalIgnoreCase)) Then Return
        tokens.Add(token)
    End Sub

    Private Shared Function IsWeakQueryToken(token As String) As Boolean
        Select Case token
            Case "给我", "帮我", "我要", "请帮", "一封", "写一", "生成", "创建", "实现", "当前", "文档"
                Return True
        End Select
        Return token.Length < 2
    End Function

    Private Shared Function ContainsCjk(value As String) As Boolean
        If String.IsNullOrEmpty(value) Then Return False
        Return value.Any(Function(ch) IsCjk(ch))
    End Function

    Private Shared Function IsCjk(ch As Char) As Boolean
        Dim code = AscW(ch)
        Return (code >= &H4E00 AndAlso code <= &H9FFF) OrElse
               (code >= &H3400 AndAlso code <= &H4DBF) OrElse
               (code >= &HF900 AndAlso code <= &HFAFF)
    End Function

    Private Shared Sub AppendSkillScriptInfo(skillParts As List(Of String), skill As SkillFileDefinition)
        If skill Is Nothing OrElse skill.Scripts Is Nothing OrElse skill.Scripts.Count = 0 Then Return

        Dim scriptInfo As New List(Of String)()
        scriptInfo.Add("**可执行脚本：**")
        For Each script In skill.Scripts
            scriptInfo.Add($"- `{script.FileName}` ({script.ScriptType})" &
                If(Not String.IsNullOrEmpty(script.Description), $" - {script.Description}", ""))
        Next

        scriptInfo.Add("")
        scriptInfo.Add("**脚本调用格式：**")
        scriptInfo.Add("```json")
        scriptInfo.Add($"{{""command"": ""skill_script.{skill.Name}.{skill.Scripts(0).FileName}"", ""params"": {{""arg1"": ""value1""}}}}")
        scriptInfo.Add("```")
        skillParts.Add(String.Join(vbCrLf, scriptInfo))
    End Sub
End Class
