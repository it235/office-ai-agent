' ShareRibbon\Services\SkillsService.vb
' Skills服务：实现Claude Skills规范和渐进式披露
' 支持从文件系统目录读取Skills（类似Trae/Cursor模式）

Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Diagnostics
Imports Newtonsoft.Json
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
''' Skills使用统计（持久化）
''' </summary>
Public Class SkillUsageStats
    Public Property SkillName As String
    Public Property UsageCount As Integer
    Public Property LastUsedAt As DateTime?
    Public Property SuccessCount As Integer
    Public Property TotalTokens As Long
End Class

''' <summary>
''' Skills使用统计存储
''' </summary>
Public Class SkillsUsageStorage
    Public Property Skills As New Dictionary(Of String, SkillUsageStats)()
    Public Property LastUpdated As DateTime = DateTime.Now
End Class

''' <summary>
''' Skills服务：实现渐进式披露和智能匹配
''' 从Skills目录读取Claude规范的Skills文件
''' </summary>
Public Class SkillsService

    ' 同义词词库（用于语义匹配）
    Private Shared ReadOnly Synonyms As New Dictionary(Of String, List(Of String))() From {
        {"excel", New List(Of String) From {"电子表格", "spreadsheet", "xlsx", "xls"}},
        {"word", New List(Of String) From {"文档", "docx", "doc", "文字处理"}},
        {"powerpoint", New List(Of String) From {"ppt", "pptx", "演示", "幻灯片"}},
        {"数据", New List(Of String) From {"data", "dataset", "数据库"}},
        {"分析", New List(Of String) From {"analyze", "analysis", "统计"}},
        {"图表", New List(Of String) From {"chart", "graph", "可视化", "visualization"}},
        {"公式", New List(Of String) From {"function", "formula", "函数", "计算"}},
        {"格式", New List(Of String) From {"format", "样式", "style", "排版"}},
        {"表格", New List(Of String) From {"table", "range", "区域"}},
        {"单元格", New List(Of String) From {"cell", "单元格"}},
        {"脚本", New List(Of String) From {"script", "vba", "宏", "macro"}},
        {"模板", New List(Of String) From {"template", "模板"}},
        {"报告", New List(Of String) From {"report", "报表", "summary"}},
        {"清理", New List(Of String) From {"clean", "清洗", "整理"}},
        {"转换", New List(Of String) From {"convert", "transform", "转换"}},
        {"批量", New List(Of String) From {"batch", "批量", "mass"}},
        {"智能", New List(Of String) From {"ai", "智能", "smart"}},
        {"自动", New List(Of String) From {"auto", "自动", "automation"}},
        {"助手", New List(Of String) From {"assistant", "helper", "助手"}},
        {"专家", New List(Of String) From {"expert", "specialist", "专家"}},
        {"顾问", New List(Of String) From {"advisor", "consultant", "顾问"}}
    }

    ' 使用统计文件路径
    Private Shared ReadOnly UsageStatsPath As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        ConfigSettings.OfficeAiAppDataFolder,
        "skills_usage.json"
    )

    ' 缓存的使用统计
    Private Shared _usageStorage As SkillsUsageStorage = Nothing
    Private Shared _usageStorageLock As New Object()

    ''' <summary>
    ''' 获取Skills目录（用于渐进式披露的第一步）
    ''' 只返回Skill的元数据，不返回详细内容
    ''' </summary>
    Public Shared Function GetSkillsCatalog() As List(Of SkillFileDefinition)
        Dim skills = SkillsDirectoryService.GetAllSkills()

        ' 加载使用统计并合并到Skills
        LoadUsageStats()
        For Each skill In skills
            If _usageStorage.Skills.ContainsKey(skill.Name.ToLowerInvariant()) Then
                Dim stats = _usageStorage.Skills(skill.Name.ToLowerInvariant())
                skill.UsageCount = stats.UsageCount
                skill.LastUsedAt = stats.LastUsedAt
            End If
        Next

        Return skills
    End Function

    ''' <summary>
    ''' 智能匹配Skills（基于用户查询，支持语义匹配）
    ''' </summary>
    Public Shared Function MatchSkills(userQuery As String, Optional topN As Integer = 5) As List(Of SkillMatchResult)
        Dim results As New List(Of SkillMatchResult)()
        Dim allSkills = GetSkillsCatalog()

        If allSkills.Count = 0 Then
            Return results
        End If

        Dim queryLower = userQuery.ToLowerInvariant()
        Dim queryWords = TokenizeQuery(queryLower)

        For Each skill In allSkills
            Dim matchResult = CalculateMatchScore(queryLower, queryWords, skill)
            If matchResult.MatchScore > 0 Then
                results.Add(matchResult)
            End If
        Next

        ' 按匹配分数排序，分数相同则按使用次数排序
        Return results.OrderByDescending(Function(r) r.MatchScore) _
                      .ThenByDescending(Function(r) r.Skill.UsageCount) _
                      .Take(topN).ToList()
    End Function

    ''' <summary>
    ''' 将查询分词
    ''' </summary>
    Private Shared Function TokenizeQuery(query As String) As List(Of String)
        Dim words As New List(Of String)()

        ' 简单分词：按常见分隔符分割
        Dim tokens = query.Split({" "c, ","c, "，"c, "。"c, "."c, "、"c, "/"c, "\"c,
                                 "("c, ")"c, "["c, "]"c, "{"c, "}"c, "："c, ":"c,
                                 "!"c, "！"c, "?"c, "？"c, ";"c, "；"c},
                                 StringSplitOptions.RemoveEmptyEntries)

        For Each token In tokens
            Dim t = token.Trim()
            If t.Length > 0 Then
                words.Add(t)
                ' 添加同义词
                For Each kvp In Synonyms
                    If kvp.Value.Contains(t) OrElse kvp.Key = t Then
                        If Not words.Contains(kvp.Key) Then
                            words.Add(kvp.Key)
                        End If
                        For Each syn In kvp.Value
                            If Not words.Contains(syn) Then
                                words.Add(syn)
                            End If
                        Next
                    End If
                Next
            End If
        Next

        Return words.Distinct().ToList()
    End Function

    ''' <summary>
    ''' 计算Skill匹配分数（增强版，支持语义匹配）
    ''' </summary>
    Private Shared Function CalculateMatchScore(queryLower As String, queryWords As List(Of String), skill As SkillFileDefinition) As SkillMatchResult
        Dim score As Double = 0
        Dim matchedKeywords As New List(Of String)()

        ' === 1. 精确匹配名称（最高权重） ===
        If Not String.IsNullOrWhiteSpace(skill.Name) Then
            Dim nameLower = skill.Name.ToLowerInvariant()
            If queryLower.Contains(nameLower) Then
                score += 30  ' 提高名称匹配权重
                matchedKeywords.Add(skill.Name)
            End If

            ' 查询词完全包含在Skill名称中
            For Each word In queryWords
                If word.Length > 1 AndAlso nameLower.Contains(word) Then
                    score += 10
                    If Not matchedKeywords.Contains(word) Then
                        matchedKeywords.Add(word)
                    End If
                End If
            Next
        End If

        ' === 2. 匹配描述词（增强版） ===
        If Not String.IsNullOrWhiteSpace(skill.Description) Then
            Dim descLower = skill.Description.ToLowerInvariant()

            ' 查询词匹配
            For Each word In queryWords
                If word.Length > 1 AndAlso descLower.Contains(word) Then
                    score += 5
                    If Not matchedKeywords.Contains(word) Then
                        matchedKeywords.Add(word)
                    End If
                End If
            Next
        End If

        ' === 3. 匹配 tags（增强版） ===
        If skill.Tags IsNot Nothing Then
            For Each tag In skill.Tags
                If Not String.IsNullOrWhiteSpace(tag) Then
                    Dim tagLower = tag.ToLowerInvariant()

                    ' 直接匹配
                    If queryLower.Contains(tagLower) Then
                        score += 8
                        If Not matchedKeywords.Contains(tag) Then
                            matchedKeywords.Add(tag)
                        End If
                    End If

                    ' 同义词匹配
                    For Each word In queryWords
                        If Synonyms.ContainsKey(word) AndAlso Synonyms(word).Contains(tagLower) Then
                            score += 6
                            If Not matchedKeywords.Contains(tag) Then
                                matchedKeywords.Add(tag)
                            End If
                            Exit For
                        End If
                        If Synonyms.ContainsKey(tagLower) AndAlso Synonyms(tagLower).Contains(word) Then
                            score += 6
                            If Not matchedKeywords.Contains(tag) Then
                                matchedKeywords.Add(tag)
                            End If
                            Exit For
                        End If
                    Next
                End If
            Next
        End If

        ' === 4. 匹配Skill内容（深度匹配） ===
        If Not String.IsNullOrWhiteSpace(skill.Content) Then
            Dim contentLower = skill.Content.ToLowerInvariant()
            Dim matchCount As Integer = 0

            For Each word In queryWords
                If word.Length > 1 AndAlso contentLower.Contains(word) Then
                    matchCount += 1
                    If matchCount <= 3 Then  ' 最多3个词有额外加分
                        score += 3
                    End If
                End If
            Next
        End If

        ' === 5. 使用频率加权（热门 Skill 加分，最多 +5） ===
        If skill.UsageCount > 0 Then
            score += Math.Min(5.0, skill.UsageCount * 0.8)
        End If

        ' === 6. 最近使用加分（活跃Skill） ===
        If skill.LastUsedAt.HasValue Then
            Dim daysSinceUse = (DateTime.Now - skill.LastUsedAt.Value).TotalDays
            If daysSinceUse < 1 Then
                score += 5  ' 24小时内用过
            ElseIf daysSinceUse < 7 Then
                score += 3  ' 一周内用过
            ElseIf daysSinceUse < 30 Then
                score += 1  ' 一个月内用过
            End If
        End If

        Return New SkillMatchResult With {
            .Skill = skill,
            .MatchScore = score,
            .MatchedKeywords = matchedKeywords
        }
    End Function

    ' 批量保存计数器
    Private Shared _unsavedChanges As Integer = 0
    Private Const SAVE_BATCH_SIZE As Integer = 10

    ''' <summary>
    ''' 记录 Skill 使用情况（持久化，批量写入）
    ''' </summary>
    Public Shared Sub RecordSkillUsage(skillName As String, Optional success As Boolean = True, Optional tokensUsed As Long = 0)
        Try
            LoadUsageStats()

            Dim key = skillName.ToLowerInvariant()
            If Not _usageStorage.Skills.ContainsKey(key) Then
                _usageStorage.Skills(key) = New SkillUsageStats With {
                    .SkillName = skillName,
                    .UsageCount = 0,
                    .SuccessCount = 0,
                    .TotalTokens = 0
                }
            End If

            Dim stats = _usageStorage.Skills(key)
            stats.UsageCount += 1
            stats.LastUsedAt = DateTime.Now
            If success Then
                stats.SuccessCount += 1
            End If
            stats.TotalTokens += tokensUsed

            ' 批量保存：每N次或程序退出时才写文件
            _unsavedChanges += 1
            If _unsavedChanges >= SAVE_BATCH_SIZE Then
                SaveUsageStats()
                _unsavedChanges = 0
            End If

            ' 更新内存中的Skill对象（使用缓存，避免刷新开销）
            Dim skill = SkillsDirectoryService.GetAllSkills().FirstOrDefault(Function(s) String.Equals(s.Name, skillName, StringComparison.OrdinalIgnoreCase))
            If skill IsNot Nothing Then
                skill.UsageCount = stats.UsageCount
                skill.LastUsedAt = stats.LastUsedAt
            End If

            Debug.WriteLine($"[SkillsService] 记录 Skill 使用: {skillName}, 累计: {stats.UsageCount}")
        Catch ex As Exception
            Debug.WriteLine($"[SkillsService] RecordSkillUsage 失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 强制保存使用统计（程序退出时调用）
    ''' </summary>
    Public Shared Sub FlushUsageStats()
        If _unsavedChanges > 0 Then
            SaveUsageStats()
            _unsavedChanges = 0
        End If
    End Sub

    ''' <summary>
    ''' 加载使用统计
    ''' </summary>
    Private Shared Sub LoadUsageStats()
        SyncLock _usageStorageLock
            If _usageStorage IsNot Nothing Then
                Return
            End If

            Try
                If File.Exists(UsageStatsPath) Then
                    Dim json = File.ReadAllText(UsageStatsPath)
                    _usageStorage = JsonConvert.DeserializeObject(Of SkillsUsageStorage)(json)
                End If
            Catch ex As Exception
                Debug.WriteLine($"[SkillsService] LoadUsageStats 失败: {ex.Message}")
            End Try

            If _usageStorage Is Nothing Then
                _usageStorage = New SkillsUsageStorage()
            End If
        End SyncLock
    End Sub

    ''' <summary>
    ''' 保存使用统计
    ''' </summary>
    Private Shared Sub SaveUsageStats()
        SyncLock _usageStorageLock
            Try
                Dim dir = Path.GetDirectoryName(UsageStatsPath)
                If Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If

                _usageStorage.LastUpdated = DateTime.Now
                Dim json = JsonConvert.SerializeObject(_usageStorage, Formatting.Indented)
                File.WriteAllText(UsageStatsPath, json)
            Catch ex As Exception
                Debug.WriteLine($"[SkillsService] SaveUsageStats 失败: {ex.Message}")
            End Try
        End SyncLock
    End Sub

    ''' <summary>
    ''' 自动装配Skills到提示词（增强版）
    ''' 根据用户查询自动匹配并注入相关Skills
    ''' </summary>
    Public Shared Function AutoInjectSkills(userQuery As String, Optional maxSkills As Integer = 3) As String
        Dim sb As New StringBuilder()

        Try
            Dim matchedSkills = MatchSkills(userQuery, maxSkills)

            If matchedSkills.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("---")
                sb.AppendLine("## 相关技能助手")
                sb.AppendLine()
                sb.AppendLine("以下技能可能对你有帮助：")
                sb.AppendLine()

                For Each result In matchedSkills
                    Dim skill = result.Skill
                    Dim starMark = If(skill.UsageCount >= 5, " ★★★", If(skill.UsageCount >= 3, " ★★", If(skill.UsageCount >= 1, " ★", "")))

                    sb.AppendLine($"### {skill.Name}{starMark}")
                    If Not String.IsNullOrWhiteSpace(skill.Description) Then
                        sb.AppendLine($"{skill.Description}")
                    End If
                    If result.MatchedKeywords.Count > 0 Then
                        sb.AppendLine($"*相关词：{String.Join(", ", result.MatchedKeywords)}*")
                    End If
                    sb.AppendLine()

                    ' 注入Skill内容
                    If Not String.IsNullOrWhiteSpace(skill.Content) Then
                        sb.AppendLine("```skill")
                        sb.AppendLine(skill.Content)
                        sb.AppendLine("```")
                        sb.AppendLine()
                    End If
                Next

                sb.AppendLine("---")
                sb.AppendLine()
            End If
        Catch ex As Exception
            Debug.WriteLine($"[SkillsService] AutoInjectSkills 失败: {ex.Message}")
        End Try

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 构建渐进式披露的第一步：Skills目录（增强版）
    ''' </summary>
    Public Shared Function BuildSkillsCatalogMessage(skills As List(Of SkillFileDefinition)) As String
        If skills Is Nothing OrElse skills.Count = 0 Then
            Return ""
        End If

        ' 按使用次数排序
        Dim sortedSkills = skills.OrderByDescending(Function(s) s.UsageCount) _
                                 .ThenByDescending(Function(s) s.LastUsedAt.GetValueOrDefault()) _
                                 .ToList()

        Dim sb As New StringBuilder()
        sb.AppendLine("## 可用的Skills（目录）")
        sb.AppendLine()
        sb.AppendLine("以下是可用的Skills，你可以根据需要选择使用：")
        sb.AppendLine()

        ' 分类显示：热门、最近使用、其他
        Dim hotSkills = sortedSkills.Where(Function(s) s.UsageCount >= 3).ToList()
        Dim recentSkills = sortedSkills.Where(Function(s) s.UsageCount < 3 AndAlso s.LastUsedAt.HasValue AndAlso (DateTime.Now - s.LastUsedAt.Value).TotalDays < 7).ToList()
        Dim otherSkills = sortedSkills.Where(Function(s) Not hotSkills.Contains(s) AndAlso Not recentSkills.Contains(s)).ToList()

        If hotSkills.Count > 0 Then
            sb.AppendLine("### 🔥 热门技能")
            sb.AppendLine()
            For Each skill In hotSkills
                sb.AppendLine($"- **{skill.Name}**")
                If Not String.IsNullOrWhiteSpace(skill.Description) Then
                    sb.AppendLine($"  {skill.Description}")
                End If
                If skill.Tags IsNot Nothing AndAlso skill.Tags.Count > 0 Then
                    sb.AppendLine($"  *标签：{String.Join(", ", skill.Tags)}*")
                End If
                sb.AppendLine($"  *使用 {skill.UsageCount} 次*")
                sb.AppendLine()
            Next
        End If

        If recentSkills.Count > 0 Then
            sb.AppendLine("### ⏰ 最近使用")
            sb.AppendLine()
            For Each skill In recentSkills
                sb.AppendLine($"- **{skill.Name}**")
                If Not String.IsNullOrWhiteSpace(skill.Description) Then
                    sb.AppendLine($"  {skill.Description}")
                End If
                sb.AppendLine()
            Next
        End If

        If otherSkills.Count > 0 Then
            sb.AppendLine("### 📚 所有技能")
            sb.AppendLine()
            For Each skill In otherSkills
                sb.AppendLine($"- **{skill.Name}**")
                If Not String.IsNullOrWhiteSpace(skill.Description) Then
                    sb.AppendLine($"  {skill.Description}")
                End If
                If skill.Tags IsNot Nothing AndAlso skill.Tags.Count > 0 Then
                    sb.AppendLine($"  *标签：{String.Join(", ", skill.Tags)}*")
                End If
                sb.AppendLine()
            Next
        End If

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

        Dim sb As New StringBuilder()
        sb.AppendLine($"## Skill：{skill.Name}")
        sb.AppendLine()

        If Not String.IsNullOrWhiteSpace(skill.Description) Then
            sb.AppendLine($"**描述**：{skill.Description}")
            sb.AppendLine()
        End If

        If skill.Tags IsNot Nothing AndAlso skill.Tags.Count > 0 Then
            sb.AppendLine($"**标签**：{String.Join(", ", skill.Tags)}")
            sb.AppendLine()
        End If

        If skill.UsageCount > 0 Then
            sb.AppendLine($"**使用统计**：{skill.UsageCount} 次使用")
            If skill.LastUsedAt.HasValue Then
                sb.AppendLine($"**最后使用**：{skill.LastUsedAt.Value.ToString("yyyy-MM-dd HH:mm")}")
            End If
            sb.AppendLine()
        End If

        ' Skill详细内容
        sb.AppendLine("**Skill内容**：")
        sb.AppendLine("```")
        sb.AppendLine(skill.Content)
        sb.AppendLine("```")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 获取最常用的Skills
    ''' </summary>
    Public Shared Function GetTopSkills(count As Integer) As List(Of SkillFileDefinition)
        Return GetSkillsCatalog() _
            .OrderByDescending(Function(s) s.UsageCount) _
            .ThenByDescending(Function(s) s.LastUsedAt.GetValueOrDefault()) _
            .Take(count) _
            .ToList()
    End Function

End Class

''' <summary>
''' 增强版Skills匹配结果
''' </summary>
Public Class SkillMatchResultEnhanced
    Public Property Skill As SkillFileDefinition
    Public Property BaseScore As Double
    Public Property TotalScore As Double
    Public Property ScoreComponents As Dictionary(Of String, Double)
    Public Property MatchedKeywords As List(Of String)
    Public Property Explanation As String  ' AI生成的推荐理由
End Class

''' <summary>
''' 上下文信息（用于增强匹配）
''' </summary>
Public Class ContextInfo
    Public Property ApplicationType As String
    Public Property CurrentTask As String
    Public Property RecentSkillsUsed As List(Of String)
    Public Property DocumentType As String
End Class

''' <summary>
''' 增强版Skills服务
''' </summary>
Public Class EnhancedSkillsService

    ''' <summary>
    ''' 智能匹配Skills（考虑用户画像、使用历史、上下文）
    ''' </summary>
    Public Shared Function MatchSkillsEnhanced(
        userQuery As String,
        Optional contextInfo As ContextInfo = Nothing,
        Optional topN As Integer = 5) As List(Of SkillMatchResultEnhanced)

        Dim results As New List(Of SkillMatchResultEnhanced)()
        Dim allSkills = SkillsService.GetSkillsCatalog()

        If allSkills.Count = 0 Then
            Return results
        End If

        ' 1. 获取用户画像
        Dim userProfile = MemoryRepository.GetAllUserProfile()

        ' 2. 基础匹配（复用原有逻辑）
        Dim baseMatches = SkillsService.MatchSkills(userQuery, topN * 2)

        ' 3. 个性化加分
        For Each baseMatch In baseMatches
            Dim skill = baseMatch.Skill
            Dim enhanced As New SkillMatchResultEnhanced()
            enhanced.Skill = skill
            enhanced.BaseScore = baseMatch.MatchScore
            enhanced.MatchedKeywords = baseMatch.MatchedKeywords

            ' 加载使用统计
            LoadSkillUsageStats(skill)

            ' 计算各维度加分
            Dim preferenceBoost = CalculatePreferenceBoost(skill, userProfile, contextInfo)
            Dim usageBoost = CalculateUsageBoost(skill)
            Dim contextBoost = CalculateContextBoost(skill, contextInfo)

            ' 综合得分
            enhanced.TotalScore = baseMatch.MatchScore * 0.4 +
                                preferenceBoost * 0.25 +
                                usageBoost * 0.2 +
                                contextBoost * 0.15

            enhanced.ScoreComponents = New Dictionary(Of String, Double) From {
                {"base", baseMatch.MatchScore},
                {"preference", preferenceBoost},
                {"usage", usageBoost},
                {"context", contextBoost}
            }

            results.Add(enhanced)
        Next

        ' 4. 排序并返回
        Return results.OrderByDescending(Function(m) m.TotalScore).Take(topN).ToList()
    End Function

    ''' <summary>
    ''' 加载技能使用统计
    ''' </summary>
    Private Shared Sub LoadSkillUsageStats(skill As SkillFileDefinition)
        Try
            Dim usage = MemoryRepository.GetSkillUsage(skill.Name)
            If usage IsNot Nothing Then
                skill.UsageCount = usage.UsageCount
                If Not String.IsNullOrWhiteSpace(usage.LastUsedAt) Then
                    Dim dt As DateTime
                    If DateTime.TryParse(usage.LastUsedAt, dt) Then
                        skill.LastUsedAt = dt
                    End If
                End If
                If usage.UsageCount > 0 Then
                    skill.SuccessRate = CDbl(usage.SuccessCount) / usage.UsageCount
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"[EnhancedSkillsService] 加载技能使用统计失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 计算用户偏好加分
    ''' </summary>
    Private Shared Function CalculatePreferenceBoost(
        skill As SkillFileDefinition,
        userProfile As Dictionary(Of String, UserProfileItem),
        contextInfo As ContextInfo) As Double

        Dim boost As Double = 0

        ' 检查用户使用过的应用
        If userProfile.ContainsKey("preferred_application") Then
            Dim preferredApp = userProfile("preferred_application").Value.ToLowerInvariant()
            Dim skillApp = If(String.IsNullOrWhiteSpace(skill.Application), "", skill.Application.ToLowerInvariant())

            ' 检查标签匹配
            Dim hasMatchingTag = skill.Tags?.Any(Function(t) t.ToLowerInvariant().Contains(preferredApp))
            If skillApp = preferredApp OrElse hasMatchingTag.GetValueOrDefault() Then
                boost += 0.15
            End If
        End If

        ' 检查领域偏好
        If userProfile.ContainsKey("domains") Then
            Dim domains = userProfile("domains").Value.Split(","c)
            For Each domain In domains
                If skill.Tags?.Any(Function(t) t.ToLowerInvariant().Contains(domain.Trim().ToLowerInvariant())) Then
                    boost += 0.1
                    Exit For
                End If
            Next
        End If

        Return Math.Min(0.3, boost)
    End Function

    ''' <summary>
    ''' 计算使用历史加分（考虑成功率、最近使用）
    ''' </summary>
    Private Shared Function CalculateUsageBoost(skill As SkillFileDefinition) As Double
        Dim boost As Double = 0

        If skill.UsageCount = 0 Then
            Return 0
        End If

        ' 基础使用量加分（最多+0.15）
        boost += Math.Min(0.15, skill.UsageCount * 0.02)

        ' 最近使用加分
        If skill.LastUsedAt.HasValue Then
            Dim daysSince = (DateTime.Now - skill.LastUsedAt.Value).TotalDays
            If daysSince < 1 Then
                boost += 0.1  ' 24小时内
            ElseIf daysSince < 7 Then
                boost += 0.05  ' 一周内
            End If
        End If

        ' 成功率加分
        If skill.SuccessRate.HasValue Then
            boost += skill.SuccessRate.Value * 0.1
        End If

        Return Math.Min(0.3, boost)
    End Function

    ''' <summary>
    ''' 计算上下文相关加分
    ''' </summary>
    Private Shared Function CalculateContextBoost(skill As SkillFileDefinition, contextInfo As ContextInfo) As Double
        If contextInfo Is Nothing Then
            Return 0
        End If

        Dim boost As Double = 0

        ' 应用类型匹配
        If Not String.IsNullOrWhiteSpace(contextInfo.ApplicationType) Then
            Dim appType = contextInfo.ApplicationType.ToLowerInvariant()
            If skill.Tags?.Any(Function(t) t.ToLowerInvariant().Contains(appType)) Then
                boost += 0.1
            End If
        End If

        ' 最近使用的技能关联
        If contextInfo.RecentSkillsUsed?.Count > 0 Then
            For Each recentSkill In contextInfo.RecentSkillsUsed
                If skill.Name.ToLowerInvariant().Contains(recentSkill.ToLowerInvariant()) OrElse
                   skill.Tags?.Any(Function(t) recentSkill.ToLowerInvariant().Contains(t.ToLowerInvariant())) Then
                    boost += 0.08
                    Exit For
                End If
            Next
        End If

        ' 当前任务关键词匹配
        If Not String.IsNullOrWhiteSpace(contextInfo.CurrentTask) Then
            Dim taskLower = contextInfo.CurrentTask.ToLowerInvariant()
            If Not String.IsNullOrWhiteSpace(skill.Description) Then
                Dim descLower = skill.Description.ToLowerInvariant()
                If descLower.Contains(taskLower) OrElse taskLower.Contains(descLower) Then
                    boost += 0.12
                End If
            End If
        End If

        Return Math.Min(0.3, boost)
    End Function

    ''' <summary>
    ''' 记录Skill使用反馈（用于持续优化）
    ''' </summary>
    Public Shared Sub RecordSkillFeedback(
        skillName As String,
        success As Boolean,
        Optional userRating As Integer? = Nothing,
        Optional tokensUsed As Long = 0)

        ' 更新使用统计
        MemoryRepository.RecordSkillUsage(skillName, success, tokensUsed)

        ' 记录详细反馈到记忆
        If userRating.HasValue Then
            Dim content = $"Skill '{skillName}' 被评为 {userRating.Value} 星，成功: {success}"
            MemoryRepository.InsertMemory(
                content,
                Nothing,
                Nothing,
                Nothing,
                "skill_feedback",
                importance:=If(userRating.Value >= 4, 0.7, 0.4),
                sourceType:="skill_feedback"
            )
        End If

        Debug.WriteLine($"[EnhancedSkillsService] 已记录技能反馈: {skillName}, 成功: {success}")
    End Sub

End Class
