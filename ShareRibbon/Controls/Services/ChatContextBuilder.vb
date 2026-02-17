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

        ' [1] 场景指令 + Skills（从 prompt_template）
        Dim systemPromptFromDb = PromptTemplateRepository.GetSystemPrompt(scenarioNorm)
        Dim skills = PromptTemplateRepository.GetSkillsForApp(scenarioNorm, appNorm)
        Dim layer1Parts As New List(Of String)()
        If Not String.IsNullOrWhiteSpace(systemPromptFromDb) Then
            layer1Parts.Add(PromptTemplateRepository.ReplaceVariables(systemPromptFromDb, vars))
        End If
        For Each sk In skills
            Dim content = PromptTemplateRepository.ReplaceVariables(sk.Content, vars)
            If Not String.IsNullOrWhiteSpace(content) Then
                layer1Parts.Add(content)
            End If
        Next
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
            Dim memoryParts As New List(Of String)()
            Dim userProfile = MemoryService.GetUserProfile()
            If Not String.IsNullOrWhiteSpace(userProfile) Then
                memoryParts.Add("[用户画像]" & vbCrLf & userProfile)
            End If
            Dim memories = MemoryService.GetRelevantMemories(currentQuery, Nothing)
            If memories IsNot Nothing AndAlso memories.Count > 0 Then
                memoryParts.Add("[相关记忆]")
                For Each m In memories
                    memoryParts.Add("- " & m.Content)
                Next
            End If
            Dim summaries = MemoryService.GetRecentSessionSummaries(Nothing)
            If summaries IsNot Nothing AndAlso summaries.Count > 0 Then
                memoryParts.Add("[近期会话]")
                For Each s In summaries
                    memoryParts.Add($"- {s.Title}: {s.Snippet}")
                Next
            End If
            If memoryParts.Count > 0 Then
                Dim memoryBlock = String.Join(vbCrLf, memoryParts)
                If result.Count > 0 AndAlso result(0).role = "system" Then
                    result(0).content = result(0).content & vbCrLf & vbCrLf & memoryBlock
                Else
                    result.Insert(0, New HistoryMessage With {.role = "system", .content = memoryBlock})
                End If
            End If
        End If

        ' [5] 当前会话滚动窗口（不含 system，只 user/assistant）
        If sessionMessages IsNot Nothing Then
            For Each msg In sessionMessages
                If msg.role <> "system" AndAlso Not String.IsNullOrEmpty(msg.content) Then
                    result.Add(New HistoryMessage With {.role = msg.role, .content = msg.content})
                End If
            Next
        End If

        ' [6] 本条 user 消息
        If Not String.IsNullOrWhiteSpace(latestUserMessage) Then
            result.Add(New HistoryMessage With {.role = "user", .content = latestUserMessage})
        End If

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
