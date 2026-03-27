' ShareRibbon\Controls\Services\ChatContextBuilder.vb
' 分层上下文组装：[0]～[6]

Imports System.Collections.Generic

''' <summary>
''' Chat 上下文构建器：按 roadmap 2.5 分层组装消息
''' </summary>
Public Class ChatContextBuilder

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

            Dim matchedSkills = SkillsService.MatchSkills(currentQuery, 5)
            If matchedSkills.Count > 0 Then
                Dim topSkill = matchedSkills.First()
                If topSkill.MatchScore >= 10 Then
                    Dim detailMessage = SkillsService.BuildSkillDetailMessage(topSkill.Skill)
                    If Not String.IsNullOrWhiteSpace(detailMessage) Then
                        skillParts.Add("#### 推荐技能（基于当前查询）")
                        skillParts.Add(detailMessage)
                    End If

                    Dim metaHints As New List(Of String)()
                    metaHints.Add($"当前推荐: {topSkill.Skill.Name}")
                    If topSkill.Skill.Tags IsNot Nothing AndAlso topSkill.Skill.Tags.Count > 0 Then
                        metaHints.Add($"标签: {String.Join(", ", topSkill.Skill.Tags)}")
                    End If
                    If Not String.IsNullOrWhiteSpace(topSkill.Skill.Compatibility) Then
                        metaHints.Add($"兼容性: {topSkill.Skill.Compatibility}")
                    End If
                    If topSkill.MatchedKeywords.Count > 0 Then
                        metaHints.Add($"匹配关键词: {String.Join(", ", topSkill.MatchedKeywords)}")
                    End If
                    skillParts.Add("> " & String.Join(" | ", metaHints))

                    Debug.WriteLine($"[ChatContextBuilder] 匹配到Skill: {topSkill.Skill.Name}, 分数: {topSkill.MatchScore:F1}, 关键词: {String.Join(", ", topSkill.MatchedKeywords)}")
                    SkillsService.RecordSkillUsage(topSkill.Skill.Name)
                End If
            Else
                Debug.WriteLine($"[ChatContextBuilder] 未匹配到Skills，提供 {skillsCatalog.Count} 个Skill目录")
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
                memParts.Add("#### 用户画像" & vbCrLf & userProfile.Trim())
            End If

            Dim memories = MemoryService.GetRelevantMemories(currentQuery, Nothing, Nothing, Nothing, appNorm)
            If memories IsNot Nothing AndAlso memories.Count > 0 Then
                ragCountOut = memories.Count
                Debug.WriteLine($"[ChatContextBuilder] 找到 {memories.Count} 条相关记忆")
                Dim memLines As New List(Of String)()
                memLines.Add("#### 相关记忆")
                For Each m In memories
                    memLines.Add("- " & m.Content)
                Next
                memParts.Add(String.Join(vbCrLf, memLines))
            Else
                Debug.WriteLine($"[ChatContextBuilder] 没有找到相关记忆，查询: {currentQuery.Substring(0, Math.Min(100, currentQuery.Length))}...")
            End If

            Dim summaries = MemoryService.GetRecentSessionSummaries(Nothing)
            If summaries IsNot Nothing AndAlso summaries.Count > 0 Then
                Debug.WriteLine($"[ChatContextBuilder] 找到 {summaries.Count} 条近期会话")
                Dim sumLines As New List(Of String)()
                sumLines.Add("#### 近期会话")
                For Each s In summaries
                    sumLines.Add($"- {s.Title}: {s.Snippet}")
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

    ''' <summary>
    ''' 简化：仅注入 Memory 层到现有 system，用于增量集成
    ''' </summary>
    ''' <param name="enableMemory">为 False 时直接返回 baseSystem</param>
    Public Shared Function AppendMemoryToSystemPrompt(baseSystem As String, currentQuery As String, Optional enableMemory As Boolean = True, Optional appType As String = Nothing) As String
        If Not enableMemory Then Return baseSystem

        Dim parts As New List(Of String)()
        If Not String.IsNullOrWhiteSpace(baseSystem) Then parts.Add(baseSystem)

        Dim userProfile = MemoryService.GetUserProfile()
        If Not String.IsNullOrWhiteSpace(userProfile) Then
            parts.Add("[用户画像]" & vbCrLf & userProfile)
        End If
        Dim memories = MemoryService.GetRelevantMemories(currentQuery, Nothing, Nothing, Nothing, appType)
        If memories IsNot Nothing AndAlso memories.Count > 0 Then
            parts.Add("[相关记忆]")
            For Each m In memories
                parts.Add("- " & m.Content)
            Next
        End If
        Dim summaries = MemoryService.GetRecentSessionSummaries(Nothing)
        If summaries IsNot Nothing AndAlso summaries.Count > 0 Then
            parts.Add("[近期会话]")
            For Each s In summaries
                parts.Add($"- {s.Title}: {s.Snippet}")
            Next
        End If

        If parts.Count <= 1 Then Return baseSystem
        Return String.Join(vbCrLf & vbCrLf, parts)
    End Function
End Class
