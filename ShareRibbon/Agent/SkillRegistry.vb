Imports System.IO
Imports System.Linq
Imports Newtonsoft.Json

Namespace Agent

    ''' <summary>
    ''' Agent 技能定义
    ''' </summary>
    Public Class AgentSkill
        Public Property Id As String
        Public Property Name As String
        Public Property Description As String
        Public Property TriggerPatterns As New List(Of String)()
        Public Property RequiredTools As New List(Of String)()
        Public Property PromptTemplate As String
        Public Property MaxSteps As Integer = 8
        Public Property AutoApprove As Boolean = False
    End Class

    ''' <summary>
    ''' 技能注册表
    ''' </summary>
    Public Class SkillRegistry
        Private ReadOnly _skills As New Dictionary(Of String, AgentSkill)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>
        ''' 从目录加载技能定义
        ''' </summary>
        Public Sub LoadFromDirectory(dir As String)
            If Not Directory.Exists(dir) Then Return
            For Each file In Directory.GetFiles(dir, "*.json")
                Try
                    Dim skill = JsonConvert.DeserializeObject(Of AgentSkill)(System.IO.File.ReadAllText(file))
                    If skill IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(skill.Id) Then
                        _skills(skill.Id) = skill
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"[SkillRegistry] 加载技能失败 {file}: {ex.Message}")
                End Try
            Next
        End Sub

        ''' <summary>
        ''' 注册技能
        ''' </summary>
        Public Sub RegisterSkill(skill As AgentSkill)
            _skills(skill.Id) = skill
        End Sub

        ''' <summary>
        ''' 根据用户输入匹配最相关的技能
        ''' 返回匹配度最高的技能，无匹配返回 Nothing
        ''' </summary>
        Public Function MatchSkill(userInput As String) As AgentSkill
            If String.IsNullOrWhiteSpace(userInput) Then Return Nothing

            Dim inputLower = userInput.ToLower()
            Dim bestMatch As AgentSkill = Nothing
            Dim bestScore As Integer = 0

            For Each skill In _skills.Values
                Dim score = 0

                ' 关键词匹配
                If skill.TriggerPatterns IsNot Nothing Then
                    For Each pattern In skill.TriggerPatterns
                        If inputLower.Contains(pattern.ToLower()) Then
                            score += 2
                        End If
                    Next
                End If

                ' 描述匹配
                If skill.Description > "" AndAlso inputLower.Contains(skill.Description.ToLower()) Then
                    score += 1
                End If

                ' 名称匹配
                If skill.Name > "" AndAlso inputLower.Contains(skill.Name.ToLower()) Then
                    score += 3
                End If

                If score > bestScore Then
                    bestScore = score
                    bestMatch = skill
                End If
            Next

            ' 至少需要匹配一个关键词
            If bestScore >= 2 Then
                Return bestMatch
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' 获取所有技能
        ''' </summary>
        Public Function GetAllSkills() As List(Of AgentSkill)
            Return _skills.Values.ToList()
        End Function

        ''' <summary>
        ''' 获取技能
        ''' </summary>
        Public Function GetSkill(id As String) As AgentSkill
            If _skills.ContainsKey(id) Then Return _skills(id)
            Return Nothing
        End Function

        ''' <summary>
        ''' 技能数量
        ''' </summary>
        Public ReadOnly Property SkillCount As Integer
            Get
                Return _skills.Count
            End Get
        End Property
    End Class

End Namespace
