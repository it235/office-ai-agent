Imports System.IO
Imports System.Diagnostics
Imports Newtonsoft.Json

''' <summary>
''' Ralph Loop 记忆系统 - 持久化存储任务执行历史和上下文
''' </summary>
Public Class RalphLoopMemory
    Private Shared ReadOnly MemoryFilePath As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "OfficeAiAgent",
        "ralph_memory.json"
    )

    ' 单例模式
    Private Shared _instance As RalphLoopMemory
    Private Shared ReadOnly _lock As New Object()

    Public Shared ReadOnly Property Instance As RalphLoopMemory
        Get
            If _instance Is Nothing Then
                SyncLock _lock
                    If _instance Is Nothing Then
                        _instance = New RalphLoopMemory()
                    End If
                End SyncLock
            End If
            Return _instance
        End Get
    End Property

    ' 记忆数据
    Public Property MemoryData As RalphMemoryData

    Private Sub New()
        Load()
    End Sub

    ''' <summary>
    ''' 加载记忆数据
    ''' </summary>
    Public Sub Load()
        Try
            If File.Exists(MemoryFilePath) Then
                Dim json = File.ReadAllText(MemoryFilePath)
                MemoryData = JsonConvert.DeserializeObject(Of RalphMemoryData)(json)
            End If
        Catch ex As Exception
            Debug.WriteLine($"[RalphLoopMemory] 加载失败: {ex.Message}")
        End Try

        If MemoryData Is Nothing Then
            MemoryData = New RalphMemoryData()
        End If

        ' 确保集合初始化
        If MemoryData.TaskHistory Is Nothing Then MemoryData.TaskHistory = New List(Of RalphTaskRecord)()
        If MemoryData.LongTermMemory Is Nothing Then MemoryData.LongTermMemory = New Dictionary(Of String, String)()
        If MemoryData.TaskTemplates Is Nothing Then MemoryData.TaskTemplates = New List(Of TaskTemplate)()
        If MemoryData.SavedSessions Is Nothing Then MemoryData.SavedSessions = New List(Of RalphLoopSession)()
    End Sub

    ''' <summary>
    ''' 保存记忆数据
    ''' </summary>
    Public Sub Save()
        Try
            Dim dir = Path.GetDirectoryName(MemoryFilePath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If
            Dim json = JsonConvert.SerializeObject(MemoryData, Formatting.Indented)
            File.WriteAllText(MemoryFilePath, json)
        Catch ex As Exception
            Debug.WriteLine($"[RalphLoopMemory] 保存失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 添加任务记录
    ''' </summary>
    Public Sub AddTaskRecord(task As RalphTaskRecord)
        MemoryData.TaskHistory.Add(task)
        ' 保持历史记录不超过100条
        If MemoryData.TaskHistory.Count > 100 Then
            MemoryData.TaskHistory.RemoveAt(0)
        End If
        Save()
    End Sub

    ''' <summary>
    ''' 获取当前活动循环
    ''' </summary>
    Public Function GetActiveLoop() As RalphLoopSession
        Return MemoryData.ActiveLoop
    End Function

    ''' <summary>
    ''' 设置当前活动循环
    ''' </summary>
    Public Sub SetActiveLoop(loopSession As RalphLoopSession)
        MemoryData.ActiveLoop = loopSession
        Save()
    End Sub

    ''' <summary>
    ''' 清除当前活动循环
    ''' </summary>
    Public Sub ClearActiveLoop()
        MemoryData.ActiveLoop = Nothing
        Save()
    End Sub

    ''' <summary>
    ''' 添加知识点到长期记忆
    ''' </summary>
    Public Sub AddKnowledge(key As String, value As String)
        MemoryData.LongTermMemory(key) = value
        Save()
    End Sub

    ''' <summary>
    ''' 获取相关知识（简单关键词匹配）
    ''' </summary>
    Public Function GetRelevantKnowledge(query As String) As List(Of String)
        Dim results As New List(Of String)
        If MemoryData.LongTermMemory Is Nothing OrElse MemoryData.LongTermMemory.Count = 0 Then
            Return results
        End If

        Dim keywords = query.ToLower().Split({" "c, "，"c, ","c}, StringSplitOptions.RemoveEmptyEntries)

        For Each kvp In MemoryData.LongTermMemory
            For Each keyword In keywords
                If keyword.Length > 1 AndAlso (kvp.Key.ToLower().Contains(keyword) OrElse kvp.Value.ToLower().Contains(keyword)) Then
                    results.Add($"{kvp.Key}: {kvp.Value}")
                    Exit For
                End If
            Next
        Next

        Return results
    End Function

    ''' <summary>
    ''' 保存当前会话为模板
    ''' </summary>
    Public Sub SaveAsTemplate(templateName As String, Optional description As String = "")
        If MemoryData.ActiveLoop Is Nothing Then
            Return
        End If

        Dim template As New TaskTemplate()
        template.Id = Guid.NewGuid().ToString()
        template.Name = templateName
        template.Description = description
        template.OriginalGoal = MemoryData.ActiveLoop.OriginalGoal
        template.ApplicationType = MemoryData.ActiveLoop.ApplicationType
        template.CreatedAt = DateTime.Now
        template.Steps = MemoryData.ActiveLoop.Steps.Select(Function(s) New TemplateStep() With {
            .StepNumber = s.StepNumber,
            .Description = s.Description,
            .Intent = s.Intent,
            .RollbackHint = s.RollbackHint
        }).ToList()

        MemoryData.TaskTemplates.Add(template)
        Save()
    End Sub

    ''' <summary>
    ''' 获取任务模板
    ''' </summary>
    Public Function GetTemplates(Optional appType As String = "") As List(Of TaskTemplate)
        If String.IsNullOrWhiteSpace(appType) Then
            Return MemoryData.TaskTemplates.ToList()
        End If
        Return MemoryData.TaskTemplates.Where(Function(t) t.ApplicationType = appType).ToList()
    End Function

    ''' <summary>
    ''' 从模板创建会话
    ''' </summary>
    Public Function CreateSessionFromTemplate(templateId As String, newGoal As String) As RalphLoopSession
        Dim template = MemoryData.TaskTemplates.FirstOrDefault(Function(t) t.Id = templateId)
        If template Is Nothing Then
            Return Nothing
        End If

        Dim session As New RalphLoopSession()
        session.Id = Guid.NewGuid().ToString()
        session.StartTime = DateTime.Now
        session.OriginalGoal = If(String.IsNullOrWhiteSpace(newGoal), template.OriginalGoal, newGoal)
        session.ApplicationType = template.ApplicationType
        session.Status = RalphLoopStatus.Ready

        For Each tStep In template.Steps
            session.Steps.Add(New RalphLoopStep() With {
                .StepNumber = tStep.StepNumber,
                .Description = tStep.Description,
                .Intent = tStep.Intent,
                .Status = RalphStepStatus.Pending,
                .RollbackHint = tStep.RollbackHint
            })
        Next

        session.TotalSteps = session.Steps.Count
        Return session
    End Function

    ''' <summary>
    ''' 查找相似历史任务
    ''' </summary>
    Public Function FindSimilarTasks(goal As String, Optional appType As String = "", Optional maxCount As Integer = 5) As List(Of RalphTaskRecord)
        Dim results = MemoryData.TaskHistory.AsEnumerable()

        If Not String.IsNullOrWhiteSpace(appType) Then
            results = results.Where(Function(t) t.ApplicationType = appType)
        End If

        ' 简单关键词匹配
        Dim keywords = goal.ToLower().Split({" "c, "，"c, ","c}, StringSplitOptions.RemoveEmptyEntries)
        If keywords.Length > 0 Then
            results = results.Where(Function(t)
                                        For Each keyword In keywords
                                            If keyword.Length > 1 AndAlso (t.UserInput?.ToLower().Contains(keyword) OrElse t.Plan?.ToLower().Contains(keyword)) Then
                                                Return True
                                            End If
                                        Next
                                        Return False
                                    End Function)
        End If

        Return results.OrderByDescending(Function(t) t.Timestamp) _
                      .Take(maxCount) _
                      .ToList()
    End Function

    ''' <summary>
    ''' 保存会话（用于断点续传）
    ''' </summary>
    Public Sub SaveSession()
        If MemoryData.ActiveLoop Is Nothing Then
            Return
        End If

        ' 先移除已存在的同ID会话
        MemoryData.SavedSessions.RemoveAll(Function(s) s.Id = MemoryData.ActiveLoop.Id)

        ' 深拷贝
        Dim sessionJson = JsonConvert.SerializeObject(MemoryData.ActiveLoop)
        Dim savedSession = JsonConvert.DeserializeObject(Of RalphLoopSession)(sessionJson)
        savedSession.SavedAt = DateTime.Now

        MemoryData.SavedSessions.Add(savedSession)

        ' 最多保留10个保存的会话
        If MemoryData.SavedSessions.Count > 10 Then
            MemoryData.SavedSessions = MemoryData.SavedSessions _
                .OrderByDescending(Function(s) s.SavedAt) _
                .Take(10) _
                .ToList()
        End If

        Save()
    End Sub

    ''' <summary>
    ''' 恢复保存的会话
    ''' </summary>
    Public Function RestoreSession(sessionId As String) As RalphLoopSession
        Dim savedSession = MemoryData.SavedSessions.FirstOrDefault(Function(s) s.Id = sessionId)
        If savedSession Is Nothing Then
            Return Nothing
        End If

        ' 深拷贝
        Dim sessionJson = JsonConvert.SerializeObject(savedSession)
        Dim restoredSession = JsonConvert.DeserializeObject(Of RalphLoopSession)(sessionJson)
        restoredSession.SavedAt = Nothing

        MemoryData.ActiveLoop = restoredSession
        Save()

        Return restoredSession
    End Function

    ''' <summary>
    ''' 获取可恢复的会话列表
    ''' </summary>
    Public Function GetRecoverableSessions() As List(Of RalphLoopSession)
        Return MemoryData.SavedSessions.OrderByDescending(Function(s) s.SavedAt).ToList()
    End Function

    ''' <summary>
    ''' 删除模板
    ''' </summary>
    Public Sub DeleteTemplate(templateId As String)
        MemoryData.TaskTemplates.RemoveAll(Function(t) t.Id = templateId)
        Save()
    End Sub
End Class

''' <summary>
''' 记忆数据结构
''' </summary>
Public Class RalphMemoryData
    ''' <summary>
    ''' 任务执行历史
    ''' </summary>
    Public Property TaskHistory As New List(Of RalphTaskRecord)

    ''' <summary>
    ''' 当前活动的循环会话
    ''' </summary>
    Public Property ActiveLoop As RalphLoopSession

    ''' <summary>
    ''' 长期记忆（知识库）
    ''' </summary>
    Public Property LongTermMemory As New Dictionary(Of String, String)

    ''' <summary>
    ''' 任务模板库
    ''' </summary>
    Public Property TaskTemplates As New List(Of TaskTemplate)

    ''' <summary>
    ''' 保存的会话（断点续传）
    ''' </summary>
    Public Property SavedSessions As New List(Of RalphLoopSession)
End Class

''' <summary>
''' 任务模板
''' </summary>
Public Class TaskTemplate
    Public Property Id As String
    Public Property Name As String
    Public Property Description As String
    Public Property OriginalGoal As String
    Public Property ApplicationType As String
    Public Property CreatedAt As DateTime
    Public Property Steps As New List(Of TemplateStep)()
End Class

''' <summary>
''' 模板步骤
''' </summary>
Public Class TemplateStep
    Public Property StepNumber As Integer
    Public Property Description As String
    Public Property Intent As String
    Public Property RollbackHint As String
End Class

''' <summary>
''' 任务记录
''' </summary>
Public Class RalphTaskRecord
    Public Property Id As String = Guid.NewGuid().ToString()
    Public Property Timestamp As DateTime = DateTime.Now
    Public Property UserInput As String
    Public Property Intent As String
    Public Property Plan As String
    Public Property Result As String
    Public Property Success As Boolean
    Public Property ApplicationType As String ' Excel/Word/PowerPoint
End Class

''' <summary>
''' 循环会话
''' </summary>
Public Class RalphLoopSession
    Public Property Id As String = Guid.NewGuid().ToString()
    Public Property StartTime As DateTime = DateTime.Now
    Public Property OriginalGoal As String
    Public Property CurrentStep As Integer = 0
    Public Property TotalSteps As Integer = 0
    Public Property Steps As New List(Of RalphLoopStep)
    Public Property Status As RalphLoopStatus = RalphLoopStatus.Planning
    Public Property ApplicationType As String

    ''' <summary>
    ''' 保存时间（用于断点续传）
    ''' </summary>
    Public Property SavedAt As DateTime?

    ''' <summary>
    ''' 获取可以并行执行的步骤组
    ''' </summary>
    Public Function GetParallelExecutableSteps() As List(Of List(Of RalphLoopStep))
        Dim result As New List(Of List(Of RalphLoopStep))()

        ' 按依赖关系分组
        Dim pendingSteps = Steps.Where(Function(s) s.Status = RalphStepStatus.Pending).ToList()

        ' 简单策略：没有依赖的步骤可以并行
        ' 可以根据 DependsOn 进一步优化
        Dim noDepSteps = pendingSteps.Where(Function(s) s.DependsOn Is Nothing OrElse s.DependsOn.Count = 0).ToList()
        If noDepSteps.Count > 0 Then
            result.Add(noDepSteps)
        End If

        ' 有依赖的步骤，按依赖关系分组
        Dim depSteps = pendingSteps.Except(noDepSteps).ToList()
        For Each loopStep In depSteps
            result.Add(New List(Of RalphLoopStep) From {loopStep})
        Next

        Return result
    End Function
End Class

''' <summary>
''' 循环步骤
''' </summary>
Public Class RalphLoopStep
    Public Property StepNumber As Integer
    Public Property Description As String
    Public Property Intent As String
    Public Property Status As RalphStepStatus = RalphStepStatus.Pending
    Public Property Result As String
    Public Property ExecutedAt As DateTime?
    Public Property CompletedAt As DateTime?
    Public Property ErrorMessage As String

    ''' <summary>
    ''' 依赖的步骤编号列表
    ''' </summary>
    Public Property DependsOn As New List(Of Integer)()

    ''' <summary>
    ''' 回滚提示
    ''' </summary>
    Public Property RollbackHint As String

    ''' <summary>
    ''' 风险级别
    ''' </summary>
    Public Property RiskLevel As String = "safe"

    ''' <summary>
    ''' 预估时间
    ''' </summary>
    Public Property EstimatedTime As String

    ''' <summary>
    ''' 重试次数
    ''' </summary>
    Public Property RetryCount As Integer = 0

    ''' <summary>
    ''' 最大重试次数
    ''' </summary>
    Public Property MaxRetries As Integer = 3
End Class

''' <summary>
''' 循环状态
''' </summary>
Public Enum RalphLoopStatus
    Planning    ' 规划中
    Ready       ' 准备执行
    Running     ' 执行中
    Paused      ' 暂停（等待用户确认继续）
    Completed   ' 已完成
    Failed      ' 失败
    RollingBack ' 回滚中
End Enum

''' <summary>
''' 步骤状态
''' </summary>
Public Enum RalphStepStatus
    Pending     ' 待执行
    Running     ' 执行中
    Completed   ' 已完成
    Failed      ' 失败
    Skipped     ' 跳过
    RolledBack  ' 已回滚
End Enum
