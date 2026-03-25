Imports System.IO
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 历史会话服务：历史记录文件、会话管理、提示词模板、Skills 导入、记忆/用户画像
''' </summary>
Public Class HistorySessionService

    Private ReadOnly _executeScript As Func(Of String, Task)
    Private ReadOnly _chatStateService As ChatStateService
    Private ReadOnly _getAppType As Func(Of String)
    Private ReadOnly _invokeOnUiThread As Action(Of Action)
    Private ReadOnly _historyService As HistoryService

    Public Sub New(
        executeScript As Func(Of String, Task),
        chatStateService As ChatStateService,
        getAppType As Func(Of String),
        invokeOnUiThread As Action(Of Action))

        _executeScript = executeScript
        _chatStateService = chatStateService
        _getAppType = getAppType
        _invokeOnUiThread = invokeOnUiThread
        _historyService = New HistoryService(executeScript)
    End Sub

    Public Sub HandleGetHistoryFiles()
        _historyService.GetHistoryFiles()
    End Sub

    Public Sub HandleOpenHistoryFile(jsonDoc As JObject)
        _historyService.OpenHistoryFile(jsonDoc)
    End Sub

    ''' <summary>
    ''' 获取近期会话列表（来自 session_summary），供历史侧边栏展示
    ''' </summary>
    Public Async Sub HandleGetSessionList()
        Dim errorScript As String = Nothing
        Try
            Dim limit As Integer = 50
            Dim summaries = MemoryRepository.GetRecentSessionSummaries(limit)
            Dim list As New List(Of Object)()
            For Each s In summaries
                list.Add(New With {
                    .sessionId = s.SessionId,
                    .title = If(String.IsNullOrEmpty(s.Title), "会话", s.Title),
                    .snippet = If(String.IsNullOrEmpty(s.Snippet), "", s.Snippet),
                    .createdAt = s.CreatedAt,
                    .fileName = s.Title,
                    .fullPath = s.SessionId,
                    .lastModified = s.CreatedAt
                })
            Next
            Dim jsonResult As String = JsonConvert.SerializeObject(list)
            Await _executeScript($"setHistoryFilesList({jsonResult});")
        Catch ex As Exception
            Debug.WriteLine("HandleGetSessionList 失败: " & ex.Message)
            errorScript = "setHistoryFilesList([]);"
        End Try
        If errorScript IsNot Nothing Then
            Await _executeScript(errorScript)
        End If
    End Sub

    ''' <summary>
    ''' 加载指定会话到当前 Chat 并渲染消息
    ''' </summary>
    Public Async Sub HandleLoadSession(jsonDoc As JObject)
        Try
            Dim sessionId As String = jsonDoc("sessionId")?.ToString()
            If String.IsNullOrEmpty(sessionId) Then Return
            _chatStateService.SwitchToSession(sessionId)
            Dim messages As New List(Of Object)()
            For Each m In _chatStateService.HistoryMessages
                If m.role = "user" OrElse m.role = "assistant" Then
                    messages.Add(New With {.role = m.role, .content = m.content, .createTime = m.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")})
                End If
            Next
            Dim jsonResult As String = JsonConvert.SerializeObject(messages)
            Await _executeScript($"setChatMessages({jsonResult});")
        Catch ex As Exception
            Debug.WriteLine("HandleLoadSession 失败: " & ex.Message)
            GlobalStatusStrip.ShowWarning("加载会话失败")
        End Try
    End Sub

    ''' <summary>
    ''' 新建会话：清空状态并清空聊天区域
    ''' </summary>
    Public Async Sub HandleNewSession()
        Try
            _chatStateService.StartNewSession()
            Await _executeScript("if(typeof clearChatContent==='function')clearChatContent();")
            GlobalStatusStrip.ShowInfo("已新建会话")
        Catch ex As Exception
            Debug.WriteLine("HandleNewSession 失败: " & ex.Message)
        End Try
    End Sub

    Public Async Sub HandleGetPromptTemplates(jsonDoc As JObject)
        Dim errorScript As String = Nothing
        Try
            Dim scenario As String = jsonDoc("scenario")?.ToString()
            If String.IsNullOrEmpty(scenario) Then scenario = "excel"
            Dim list = PromptTemplateRepository.ListByScenario(scenario)
            Dim arr As New List(Of Object)()
            For Each r In list
                arr.Add(New With {
                    .id = r.Id,
                    .templateName = r.TemplateName,
                    .scenario = r.Scenario,
                    .content = r.Content,
                    .isSkill = r.IsSkill,
                    .extraJson = r.ExtraJson,
                    .sort = r.Sort
                })
            Next
            Dim json = JsonConvert.SerializeObject(arr)
            Await _executeScript($"setPromptTemplatesList({json});")
        Catch ex As Exception
            Debug.WriteLine("HandleGetPromptTemplates 失败: " & ex.Message)
            errorScript = "setPromptTemplatesList([]);"
        End Try
        If errorScript IsNot Nothing Then
            Await _executeScript(errorScript)
        End If
    End Sub

    Public Sub HandleSavePromptTemplate(jsonDoc As JObject)
        Try
            Dim id As Long = If(jsonDoc("id")?.Value(Of Long)(), 0)
            Dim templateName As String = jsonDoc("templateName")?.ToString()
            Dim scenario As String = jsonDoc("scenario")?.ToString()
            Dim content As String = jsonDoc("content")?.ToString()
            Dim isSkill As Integer = If(jsonDoc("isSkill")?.Value(Of Integer)(), 0)
            Dim extraJson As String = jsonDoc("extraJson")?.ToString()
            Dim sort As Integer = If(jsonDoc("sort")?.Value(Of Integer)(), 0)
            Dim record As New PromptTemplateRecord With {
                .Id = id,
                .TemplateName = templateName,
                .Scenario = If(String.IsNullOrEmpty(scenario), "common", scenario),
                .Content = content,
                .IsSkill = isSkill,
                .ExtraJson = If(extraJson, ""),
                .Sort = sort
            }
            If id > 0 Then
                PromptTemplateRepository.Update(record)
                GlobalStatusStrip.ShowInfo("已更新")
            Else
                PromptTemplateRepository.Insert(record)
                GlobalStatusStrip.ShowInfo("已添加")
            End If
            HandleGetPromptTemplates(JObject.Parse("{""scenario"":""" & record.Scenario & """}"))
        Catch ex As Exception
            Debug.WriteLine("HandleSavePromptTemplate 失败: " & ex.Message)
            GlobalStatusStrip.ShowWarning("保存失败: " & ex.Message)
        End Try
    End Sub

    Public Sub HandleDeletePromptTemplate(jsonDoc As JObject)
        Try
            Dim id As Long = jsonDoc("id")?.Value(Of Long)()
            If id <= 0 Then Return
            PromptTemplateRepository.Delete(id)
            GlobalStatusStrip.ShowInfo("已删除")
            Dim scenario As String = jsonDoc("scenario")?.ToString()
            If String.IsNullOrEmpty(scenario) Then scenario = "excel"
            Dim jo As JObject = JObject.FromObject(New With {.scenario = scenario})
            HandleGetPromptTemplates(jo)
        Catch ex As Exception
            Debug.WriteLine("HandleDeletePromptTemplate 失败: " & ex.Message)
            GlobalStatusStrip.ShowWarning("删除失败: " & ex.Message)
        End Try
    End Sub

    Public Async Sub HandleGetAtomicMemories(jsonDoc As JObject)
        Dim errorScript As String = Nothing
        Try
            Dim limit As Integer = If(jsonDoc("limit")?.Value(Of Integer)(), 100)
            Dim appType As String = jsonDoc("appType")?.ToString()
            If String.IsNullOrEmpty(appType) Then appType = _getAppType()
            Dim list = MemoryRepository.ListAtomicMemories(limit, 0, appType)
            Dim arr As New List(Of Object)()
            For Each r In list
                arr.Add(New With {.id = r.Id, .content = r.Content, .createTime = r.CreateTime})
            Next
            Dim json = JsonConvert.SerializeObject(arr)
            Await _executeScript($"setAtomicMemoriesList({json});")
        Catch ex As Exception
            Debug.WriteLine("HandleGetAtomicMemories 失败: " & ex.Message)
            errorScript = "setAtomicMemoriesList([]);"
        End Try
        If errorScript IsNot Nothing Then
            Await _executeScript(errorScript)
        End If
    End Sub

    Public Sub HandleDeleteAtomicMemory(jsonDoc As JObject)
        Try
            Dim id As Long = jsonDoc("id")?.Value(Of Long)()
            If id <= 0 Then Return
            MemoryRepository.DeleteAtomicMemory(id)
            GlobalStatusStrip.ShowInfo("已删除")
            Dim appType As String = jsonDoc("appType")?.ToString()
            If String.IsNullOrEmpty(appType) Then appType = _getAppType()
            Dim jo As JObject = JObject.FromObject(New With {.limit = 100, .appType = appType})
            HandleGetAtomicMemories(jo)
        Catch ex As Exception
            Debug.WriteLine("HandleDeleteAtomicMemory 失败: " & ex.Message)
            GlobalStatusStrip.ShowWarning("删除失败: " & ex.Message)
        End Try
    End Sub

    Public Async Sub HandleGetUserProfile()
        Dim errorScript As String = Nothing
        Try
            Dim content As String = MemoryRepository.GetUserProfile()
            Dim json As String = JsonConvert.SerializeObject(If(content, ""))
            Await _executeScript("setUserProfileContent(" & json & ");")
        Catch ex As Exception
            Debug.WriteLine("HandleGetUserProfile 失败: " & ex.Message)
            errorScript = "setUserProfileContent('');"
        End Try
        If errorScript IsNot Nothing Then
            Await _executeScript(errorScript)
        End If
    End Sub

    Public Sub HandleSaveUserProfile(jsonDoc As JObject)
        Try
            Dim content As String = jsonDoc("content")?.ToString()
            MemoryRepository.UpdateUserProfile(content)
            GlobalStatusStrip.ShowInfo("用户画像已保存")
        Catch ex As Exception
            Debug.WriteLine("HandleSaveUserProfile 失败: " & ex.Message)
            GlobalStatusStrip.ShowWarning("保存失败: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 从文件夹批量导入 Skill（.json/.md），与 SkillsConfigForm 的导入逻辑一致
    ''' </summary>
    Public Sub HandleImportSkillsFromFolder(jsonDoc As JObject)
        _invokeOnUiThread(Sub()
            Try
                Dim scenario As String = jsonDoc("scenario")?.ToString()
                If String.IsNullOrEmpty(scenario) Then scenario = "excel"
                Using dlg As New FolderBrowserDialog()
                    dlg.Description = "选择包含 .json / .md Skill 文件的文件夹"
                    If dlg.ShowDialog() <> DialogResult.OK Then Return
                    Dim folder = dlg.SelectedPath
                    Dim files As New List(Of String)()
                    Try
                        files.AddRange(Directory.GetFiles(folder, "*.json"))
                        files.AddRange(Directory.GetFiles(folder, "*.md"))
                    Catch ex As Exception
                        GlobalStatusStrip.ShowWarning("读取文件夹失败: " & ex.Message)
                        Return
                    End Try
                    Dim sort = 0
                    Dim count = 0
                    For Each filePath In files
                        Try
                            Dim fileContent = File.ReadAllText(filePath)
                            Dim ext = Path.GetExtension(filePath).ToLowerInvariant()
                            Dim name = Path.GetFileNameWithoutExtension(filePath)
                            Dim record As PromptTemplateRecord = Nothing
                            If ext = ".json" Then
                                Dim jo = JObject.Parse(fileContent)
                                Dim pt = jo("promptTemplate")
                                Dim ct = jo("content")
                                Dim pm = jo("prompt")
                                Dim promptTemplate = If(pt IsNot Nothing, pt.ToString(), If(ct IsNot Nothing, ct.ToString(), If(pm IsNot Nothing, pm.ToString(), "")))
                                If String.IsNullOrWhiteSpace(promptTemplate) Then Continue For
                                Dim sn = jo("skillName")
                                Dim nm = jo("name")
                                Dim skillName = If(sn IsNot Nothing, sn.ToString(), If(nm IsNot Nothing, nm.ToString(), name))
                                Dim supportedApps = If(jo("supported_apps"), jo("supportedApps"))
                                Dim extraJo As New JObject()
                                If supportedApps IsNot Nothing AndAlso TypeOf supportedApps Is JArray Then
                                    extraJo("supported_apps") = supportedApps
                                End If
                                Dim params = If(jo("parameters"), jo("params"))
                                If params IsNot Nothing Then extraJo("parameters") = params
                                Dim extra = If(extraJo.Count > 0, extraJo.ToString(), "")
                                record = New PromptTemplateRecord With {
                                    .TemplateName = skillName,
                                    .Content = promptTemplate,
                                    .IsSkill = 1,
                                    .ExtraJson = extra,
                                    .Scenario = scenario,
                                    .Sort = sort
                                }
                            Else
                                record = New PromptTemplateRecord With {
                                    .TemplateName = name,
                                    .Content = fileContent,
                                    .IsSkill = 1,
                                    .ExtraJson = "",
                                    .Scenario = scenario,
                                    .Sort = sort
                                }
                            End If
                            PromptTemplateRepository.Insert(record)
                            sort += 1
                            count += 1
                        Catch ex As Exception
                            Debug.WriteLine($"导入 {filePath} 失败: {ex.Message}")
                        End Try
                    Next
                    GlobalStatusStrip.ShowInfo($"已从文件夹导入 {count} 个 Skill")
                    Dim joRefresh As JObject = JObject.FromObject(New With {.scenario = scenario})
                    HandleGetPromptTemplates(joRefresh)
                End Using
            Catch ex As Exception
                Debug.WriteLine("HandleImportSkillsFromFolder 失败: " & ex.Message)
                GlobalStatusStrip.ShowWarning("导入失败: " & ex.Message)
            End Try
        End Sub)
    End Sub

End Class
