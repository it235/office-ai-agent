' ShareRibbon\Services\SkillsService.vb
' Skills服务：实现Claude Skills规范和渐进式披露
' 支持从文件系统目录读取Skills（类似Trae/Cursor模式）

Imports System.Collections.Generic
Imports System.Linq
Imports Newtonsoft.Json.Linq

''' <summary>
''' Skills匹配结果
''' </summary>
Public Class SkillMatchResult
    Public Property Skill As SkillFileDefinition
    Public Property MatchScore As Double
    Public Property MatchedKeywords As List(Of String)
End Class

''' <summary>
''' Skills服务：实现渐进式披露和智能匹配
''' 从Skills目录读取Claude规范的Skills文件
''' </summary>
Public Class SkillsService

    ''' <summary>
    ''' 获取Skills目录（用于渐进式披露的第一步）
    ''' 只返回Skill的元数据，不返回详细内容
    ''' </summary>
    Public Shared Function GetSkillsCatalog() As List(Of SkillFileDefinition)
        Return SkillsDirectoryService.GetAllSkills()
    End Function

    ''' <summary>
    ''' 智能匹配Skills（基于用户查询）
    ''' </summary>
    Public Shared Function MatchSkills(userQuery As String, Optional topN As Integer = 5) As List(Of SkillMatchResult)
        Dim results As New List(Of SkillMatchResult)()
        Dim allSkills = GetSkillsCatalog()

        For Each skill In allSkills
            Dim matchResult = CalculateMatchScore(userQuery, skill)
            If matchResult.MatchScore > 0 Then
                results.Add(matchResult)
            End If
        Next

        ' 按匹配分数排序
        Return results.OrderByDescending(Function(r) r.MatchScore).Take(topN).ToList()
    End Function

    ''' <summary>
    ''' 计算Skill匹配分数
    ''' </summary>
    Private Shared Function CalculateMatchScore(userQuery As String, skill As SkillFileDefinition) As SkillMatchResult
        Dim score As Double = 0
        Dim matchedKeywords As New List(Of String)()
        Dim queryLower = userQuery.ToLowerInvariant()

        '' 匹配关键词
        'If skill.Keywords IsNot Nothing AndAlso skill.Keywords.Count > 0 Then
        '    For Each keyword In skill.Keywords
        '        Dim kwLower = keyword.Trim().ToLowerInvariant()
        '        If Not String.IsNullOrWhiteSpace(kwLower) AndAlso queryLower.Contains(kwLower) Then
        '            score += 10
        '            matchedKeywords.Add(keyword.Trim())
        '        End If
        '    Next
        'End If

        ' 匹配名称
        If Not String.IsNullOrWhiteSpace(skill.Name) Then
            Dim nameLower = skill.Name.ToLowerInvariant()
            If queryLower.Contains(nameLower) Then
                score += 20
                matchedKeywords.Add(skill.Name)
            End If
        End If

        ' 匹配描述
        If Not String.IsNullOrWhiteSpace(skill.Description) Then
            Dim descLower = skill.Description.ToLowerInvariant()
            Dim descWords = descLower.Split({" "c, ","c, "，"c, "。"c, "."c}, StringSplitOptions.RemoveEmptyEntries)
            For Each word In descWords.Take(10)
                If word.Length > 2 AndAlso queryLower.Contains(word) Then
                    score += 2
                End If
            Next
        End If

        Return New SkillMatchResult With {
            .Skill = skill,
            .MatchScore = score,
            .MatchedKeywords = matchedKeywords
        }
    End Function

    ''' <summary>
    ''' 构建渐进式披露的第一步：Skills目录
    ''' 类似于"目录"，只告诉模型有哪些Skills可用
    ''' </summary>
    Public Shared Function BuildSkillsCatalogMessage(skills As List(Of SkillFileDefinition)) As String
        If skills Is Nothing OrElse skills.Count = 0 Then
            Return ""
        End If

        Dim sb As New Text.StringBuilder()
        sb.AppendLine("## 可用的Skills（目录）")
        sb.AppendLine()
        sb.AppendLine("以下是可用的Skills，你可以根据需要选择使用：")
        sb.AppendLine()

        For Each skill In skills
            sb.AppendLine($"### {skill.Name}")
            If Not String.IsNullOrWhiteSpace(skill.Description) Then
                sb.AppendLine($"- 描述：{skill.Description}")
            End If
            'If skill.Keywords IsNot Nothing AndAlso skill.Keywords.Count > 0 Then
            '    sb.AppendLine($"- 关键词：{String.Join(", ", skill.Keywords)}")
            'End If
            sb.AppendLine()
        Next

        sb.AppendLine("---")
        sb.AppendLine("**使用说明**：")
        sb.AppendLine("1. 根据用户需求，从上面的Skills中选择最相关的")
        sb.AppendLine("2. 如果需要某个Skill的详细内容，请明确指出需要哪个Skill")
        sb.AppendLine("3. 可以同时使用多个Skills")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 构建渐进式披露的第二步：披露选中的Skill详细内容
    ''' </summary>
    Public Shared Function BuildSkillDetailMessage(skill As SkillFileDefinition) As String
        If skill Is Nothing Then
            Return ""
        End If

        Dim sb As New Text.StringBuilder()
        sb.AppendLine($"## Skill：{skill.Name}")
        sb.AppendLine()

        If Not String.IsNullOrWhiteSpace(skill.Description) Then
            sb.AppendLine($"**描述**：{skill.Description}")
            sb.AppendLine()
        End If

        ' Skill详细内容
        sb.AppendLine("**Skill内容**：")
        sb.AppendLine("```")
        sb.AppendLine(skill.Content)
        sb.AppendLine("```")

        Return sb.ToString()
    End Function

End Class
