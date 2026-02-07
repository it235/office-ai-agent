Imports System.IO
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
        Dim keywords = query.ToLower().Split({" "c, "，"c, ","c}, StringSplitOptions.RemoveEmptyEntries)
        
        For Each kvp In MemoryData.LongTermMemory
            For Each keyword In keywords
                If kvp.Key.ToLower().Contains(keyword) OrElse kvp.Value.ToLower().Contains(keyword) Then
                    results.Add($"{kvp.Key}: {kvp.Value}")
                    Exit For
                End If
            Next
        Next
        
        Return results
    End Function
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
End Enum
