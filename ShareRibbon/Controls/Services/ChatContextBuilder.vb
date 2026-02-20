' ShareRibbon\Controls\Services\ChatContextBuilder.vb
' 分层上下文组装：[0]～[6]

Imports System.Collections.Generic

''' <summary>
''' Chat 上下文构建器：按 roadmap 2.5 分层组装消息
''' </summary>
Public Class ChatContextBuilder

    ''' <summary>
    ''' 构建分层消息列表
    ''' </summary>
    ''' <param name="scenario">excel/word/ppt/common</param>
    ''' <param name="appType">当前宿主类型</param>
    ''' <param name="currentQuery">用户当前输入（用于 RAG）</param>
    ''' <param name="sessionMessages">当前会话滚动窗口 (user/assistant)</param>
    ''' <param name="latestUserMessage">本条 user 消息</param>
    ''' <param name="baseSystemPrompt">已有 system 提示词（来自 PromptManager 等）</param>
    ''' <param name="variableValues">变量替换字典，如 {{选中内容}}</param>
    ''' <param name="enableMemory">是否启用 Memory（RAG、用户画像、会话摘要）</param>
    ''' <returns>按 [0]～[6] 顺序的消息列表</returns>
    Public Shared Function BuildMessages(
        scenario As String,
        appType As String,
        currentQuery As String,
        sessionMessages As List(Of HistoryMessage),
        latestUserMessage As String,
        baseSystemPrompt As String,
        variableValues As Dictionary(Of String, String),
        enableMemory As Boolean) As List(Of HistoryMessage)

        Dim result As New List(Of HistoryMessage)()
        Dim scenarioNorm = If(String.IsNullOrEmpty(scenario), "common", scenario.ToLowerInvariant())
        Dim appNorm = If(String.IsNullOrEmpty(appType), "Excel", appType)
        Dim vars = If(variableValues, New Dictionary(Of String, String)())

        ' [0] System 基础
        If Not String.IsNullOrWhiteSpace(baseSystemPrompt) Then
            result.Add(New HistoryMessage With {.role = "system", .content = baseSystemPrompt})
        End If

        ' [1] 场景指令 + Skills渐进式披露
        Dim systemPromptFromDb = PromptTemplateRepository.GetSystemPrompt(scenarioNorm)
        Dim layer1Parts As New List(Of String)()

        ' 场景系统提示词
        If Not String.IsNullOrWhiteSpace(systemPromptFromDb) Then
            layer1Parts.Add(PromptTemplateRepository.ReplaceVariables(systemPromptFromDb, vars))
        End If

        ' Skills渐进式披露：第一步 - 先提供Skills目录
        Dim skillsCatalog = SkillsService.GetSkillsCatalog()
        If skillsCatalog IsNot Nothing AndAlso skillsCatalog.Count > 0 Then
            ' 首先进行智能匹配
            Dim matchedSkills = SkillsService.MatchSkills(currentQuery, 5)

            If matchedSkills.Count > 0 Then
                ' 有匹配的Skills，先披露目录，再披露匹配的详细内容
                Dim catalogMessage = SkillsService.BuildSkillsCatalogMessage(skillsCatalog)
                If Not String.IsNullOrWhiteSpace(catalogMessage) Then
                    layer1Parts.Add(catalogMessage)
                End If

                ' 披露匹配度最高的Skill详细内容
                Dim topSkill = matchedSkills.First()
                If topSkill.MatchScore >= 10 Then
                    Dim detailMessage = SkillsService.BuildSkillDetailMessage(topSkill.Skill)
                    If Not String.IsNullOrWhiteSpace(detailMessage) Then
                        layer1Parts.Add("---")
                        layer1Parts.Add("## 推荐使用的Skill（基于你的查询）")
                        layer1Parts.Add(detailMessage)
                    End If

                    Debug.WriteLine($"[ChatContextBuilder] 匹配到Skill: {topSkill.Skill.Name}, 分数: {topSkill.MatchScore:F1}, 关键词: {String.Join(", ", topSkill.MatchedKeywords)}")
                End If
            Else
                ' 没有匹配的Skills，只提供目录让模型选择
                Dim catalogMessage = SkillsService.BuildSkillsCatalogMessage(skillsCatalog)
                If Not String.IsNullOrWhiteSpace(catalogMessage) Then
                    layer1Parts.Add(catalogMessage)
                End If
                Debug.WriteLine($"[ChatContextBuilder] 未匹配到Skills，提供 {skillsCatalog.Count} 个Skill目录")
            End If
        End If

        If layer1Parts.Count > 0 Then
            Dim layer1 = String.Join(vbCrLf & vbCrLf, layer1Parts)
            If result.Count > 0 AndAlso result(0).role = "system" Then
                result(0).content = result(0).content & vbCrLf & vbCrLf & layer1
            Else
                result.Insert(0, New HistoryMessage With {.role = "system", .content = layer1})
            End If
        End If

        ' [2] Session Metadata 可选：当前时间等
        ' 暂不注入，可后续扩展

        ' [3][4] 用户记忆 RAG + 近期会话摘要
        If enableMemory Then
            Debug.WriteLine($"[ChatContextBuilder] 启用记忆，开始检索...")
            Dim memoryParts As New List(Of String)()
            Dim userProfile = MemoryService.GetUserProfile()
            If Not String.IsNullOrWhiteSpace(userProfile) Then
                Debug.WriteLine($"[ChatContextBuilder] 找到用户画像")
                memoryParts.Add("[用户画像]" & vbCrLf & userProfile)
            Else
                Debug.WriteLine($"[ChatContextBuilder] 没有用户画像")
            End If
            Dim memories = MemoryService.GetRelevantMemories(currentQuery, Nothing)
            If memories IsNot Nothing AndAlso memories.Count > 0 Then
                Debug.WriteLine($"[ChatContextBuilder] 找到 {memories.Count} 条相关记忆")
                memoryParts.Add("[相关记忆]")
                For Each m In memories
                    Debug.WriteLine($"[ChatContextBuilder]   - {m.Content.Substring(0, Math.Min(50, m.Content.Length))}...")
                    memoryParts.Add("- " & m.Content)
                Next
            Else
                Debug.WriteLine($"[ChatContextBuilder] 没有找到相关记忆，查询内容: {currentQuery.Substring(0, Math.Min(100, currentQuery.Length))}...")
            End If
            Dim summaries = MemoryService.GetRecentSessionSummaries(Nothing)
            If summaries IsNot Nothing AndAlso summaries.Count > 0 Then
                Debug.WriteLine($"[ChatContextBuilder] 找到 {summaries.Count} 条近期会话")
                memoryParts.Add("[近期会话]")
                For Each s In summaries
                    memoryParts.Add($"- {s.Title}: {s.Snippet}")
                Next
            Else
                Debug.WriteLine($"[ChatContextBuilder] 没有近期会话")
            End If
            If memoryParts.Count > 0 Then
                Debug.WriteLine($"[ChatContextBuilder] 组装记忆块，共 {memoryParts.Count} 部分")
                Dim memoryBlock = String.Join(vbCrLf, memoryParts)
                If result.Count > 0 AndAlso result(0).role = "system" Then
                    result(0).content = result(0).content & vbCrLf & vbCrLf & memoryBlock
                Else
                    result.Insert(0, New HistoryMessage With {.role = "system", .content = memoryBlock})
                End If
            Else
                Debug.WriteLine($"[ChatContextBuilder] 没有记忆内容可注入")
            End If
        Else
            Debug.WriteLine($"[ChatContextBuilder] 记忆被禁用")
        End If

        ' [5] 当前会话滚动窗口（不含 system，只 user/assistant）
        If sessionMessages IsNot Nothing Then
            Debug.WriteLine($"[ChatContextBuilder] 添加当前会话滚动窗口，共 {sessionMessages.Count} 条原始消息")
            Dim addedCount = 0
            For Each msg In sessionMessages
                If msg.role <> "system" AndAlso Not String.IsNullOrEmpty(msg.content) Then
                    result.Add(New HistoryMessage With {.role = msg.role, .content = msg.content})
                    addedCount += 1
                    Debug.WriteLine($"[ChatContextBuilder]   - 添加 {msg.role} 消息: {msg.content.Substring(0, Math.Min(30, msg.content.Length))}...")
                End If
            Next
            Debug.WriteLine($"[ChatContextBuilder] 会话滚动窗口处理完成，共添加 {addedCount} 条消息")
        Else
            Debug.WriteLine($"[ChatContextBuilder] 没有会话消息")
        End If

        ' [6] 本条 user 消息
        If Not String.IsNullOrWhiteSpace(latestUserMessage) Then
            Debug.WriteLine($"[ChatContextBuilder] 添加本条 user 消息: {latestUserMessage.Substring(0, Math.Min(50, latestUserMessage.Length))}...")
            result.Add(New HistoryMessage With {.role = "user", .content = latestUserMessage})
        End If

        Debug.WriteLine($"[ChatContextBuilder] 构建完成，最终消息数: {result.Count}")
        Return result
    End Function

    ''' <summary>
    ''' 简化：仅注入 Memory 层到现有 system，用于增量集成
    ''' </summary>
    ''' <param name="enableMemory">为 False 时直接返回 baseSystem</param>
    Public Shared Function AppendMemoryToSystemPrompt(baseSystem As String, currentQuery As String, Optional enableMemory As Boolean = True) As String
        If Not enableMemory Then Return baseSystem

        Dim parts As New List(Of String)()
        If Not String.IsNullOrWhiteSpace(baseSystem) Then parts.Add(baseSystem)

        Dim userProfile = MemoryService.GetUserProfile()
        If Not String.IsNullOrWhiteSpace(userProfile) Then
            parts.Add("[用户画像]" & vbCrLf & userProfile)
        End If
        Dim memories = MemoryService.GetRelevantMemories(currentQuery, Nothing)
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
