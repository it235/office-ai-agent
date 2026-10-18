Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 排版服务：排版模板、样式规范、AI模板编辑器、语义排版（非Overridable方法）
''' </summary>
Public Class ReformatService

    Private ReadOnly _executeScript As Func(Of String, Task)
    Private ReadOnly _invokeOnUiThread As Action(Of Action)
    Private ReadOnly _showTemplateEditor As Func(Of ReformatTemplate, Boolean)
    Private ReadOnly _getStylePreviewCallback As Func(Of PreviewStyleCallback)
    Private ReadOnly _uploadDocxTemplateFromPath As Action(Of String)
    Private ReadOnly _textAnalyzer As Func(Of String, String, Task(Of String))

    ' 排版重试上下文
    Private _reformatRetryCount As New Dictionary(Of String, Integer)()

    ' 智能排版 v2（共享编排器实例，ChatFormatterAgent和FormattingOrchestrator使用同一实例）
    Private _chatFormatterAgent As ChatFormatterAgent = Nothing

    Private ReadOnly Property FormattingOrchestrator As SmartFormattingOrchestrator
        Get
            Return ChatFormatterAgent.Orchestrator
        End Get
    End Property

    Public ReadOnly Property ChatFormatterAgent As ChatFormatterAgent
        Get
            If _chatFormatterAgent Is Nothing Then
                _chatFormatterAgent = New ChatFormatterAgent(_executeScript, _textAnalyzer)
            End If
            Return _chatFormatterAgent
        End Get
    End Property

    Public Sub New(
        executeScript As Func(Of String, Task),
        invokeOnUiThread As Action(Of Action),
        showTemplateEditor As Func(Of ReformatTemplate, Boolean),
        getStylePreviewCallback As Func(Of PreviewStyleCallback),
        uploadDocxTemplateFromPath As Action(Of String),
        Optional textAnalyzer As Func(Of String, String, Task(Of String)) = Nothing)

        _executeScript = executeScript
        _invokeOnUiThread = invokeOnUiThread
        _showTemplateEditor = showTemplateEditor
        _getStylePreviewCallback = getStylePreviewCallback
        _uploadDocxTemplateFromPath = uploadDocxTemplateFromPath
        _textAnalyzer = textAnalyzer
    End Sub

#Region "排版模板"

    ''' <summary>
    ''' 获取排版模板列表（含docx解析出的语义映射卡片）
    ''' </summary>
    Public Sub HandleGetReformatTemplates()
        Try
            Dim templates = ReformatTemplateManager.Instance.Templates
            Dim allItems As New List(Of Object)()
            For Each t In templates
                ' 面板只渲染展示字段：缩略图Base64与源文档正文留在存储层，不进 WebView
                allItems.Add(New With {
                    .Id = t.Id,
                    .Name = t.Name,
                    .Description = t.Description,
                    .Category = t.Category,
                    .IsPreset = t.IsPreset,
                    .CreatedAt = t.CreatedAt,
                    .Layout = t.Layout,
                    .BodyStyles = t.BodyStyles,
                    .PageSettings = t.PageSettings,
                    .AiGuidance = t.AiGuidance
                })
            Next

            For Each m In SemanticMappingManager.Instance.Mappings
                If m.SourceType = SemanticMappingSourceType.FromDocxTemplate Then
                    allItems.Add(New With {
                        .Id = "docx_" & m.Id,
                        .Name = m.Name,
                        .Description = $"从Word文档提取，共{m.SemanticTags.Count}个语义标签",
                        .Category = "文档提取",
                        .IsPreset = False,
                        .IsDocxMapping = True,
                        .MappingId = m.Id,
                        .SemanticTags = m.SemanticTags,
                        .CreatedAt = m.CreatedAt
                    })
                End If
            Next

            Dim json = UtilsService.SerializeForScript(allItems)
            _executeScript($"loadReformatTemplateList({json});")
        Catch ex As Exception
            Debug.WriteLine($"HandleGetReformatTemplates 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 导入模板（含.docx/.dotx解析路由）
    ''' </summary>
    Public Sub HandleImportTemplate()
        _invokeOnUiThread(Sub()
            Try
                Dim ofd As New OpenFileDialog With {
                    .Filter = "模板文件 (*.json;*.doc;*.docx;*.dotx;*.ppt;*.pptx)|*.json;*.doc;*.docx;*.dotx;*.ppt;*.pptx|JSON文件 (*.json)|*.json|Word文档/模板 (*.doc;*.docx;*.dotx)|*.doc;*.docx;*.dotx|PowerPoint文档 (*.ppt;*.pptx)|*.ppt;*.pptx|所有文件 (*.*)|*.*",
                    .Title = "选择要导入的模板文件"
                }

                If ofd.ShowDialog() = DialogResult.OK Then
                    Dim ext = Path.GetExtension(ofd.FileName).ToLower()

                    If ext = ".docx" OrElse ext = ".dotx" Then
                        _uploadDocxTemplateFromPath(ofd.FileName)
                        Return
                    End If

                    Dim imported = ReformatTemplateManager.Instance.ImportTemplate(ofd.FileName)
                    If imported IsNot Nothing Then
                        GlobalStatusStrip.ShowInfo("模板「" & imported.Name & "」导入成功")
                        HandleGetReformatTemplates()
                    Else
                        GlobalStatusStrip.ShowWarning("模板导入失败，请检查文件格式")
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine($"HandleImportTemplate 出错: {ex.Message}")
                GlobalStatusStrip.ShowWarning($"导入模板失败: {ex.Message}")
            End Try
        End Sub)
    End Sub

    ''' <summary>
    ''' 导出模板
    ''' </summary>
    Public Sub HandleExportTemplate(jsonDoc As JObject)
        _invokeOnUiThread(Sub()
            Try
                Dim templateId = jsonDoc("templateId")?.ToString()
                Dim template = ReformatTemplateManager.Instance.GetTemplateById(templateId)

                If template Is Nothing Then
                    GlobalStatusStrip.ShowWarning("模板不存在")
                    Return
                End If

                Dim sfd As New SaveFileDialog With {
                    .Filter = "模板文件 (*.json)|*.json",
                    .Title = "导出模板",
                    .FileName = $"{template.Name}.json"
                }

                If sfd.ShowDialog() = DialogResult.OK Then
                    If ReformatTemplateManager.Instance.ExportTemplate(templateId, sfd.FileName) Then
                        GlobalStatusStrip.ShowInfo($"模板已导出到: {sfd.FileName}")
                    Else
                        GlobalStatusStrip.ShowWarning("模板导出失败")
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine($"HandleExportTemplate 出错: {ex.Message}")
                GlobalStatusStrip.ShowWarning($"导出模板失败: {ex.Message}")
            End Try
        End Sub)
    End Sub

    ''' <summary>
    ''' 复制模板
    ''' </summary>
    Public Sub HandleDuplicateTemplate(jsonDoc As JObject)
        Try
            Dim templateId = jsonDoc("templateId")?.ToString()
            Dim newName = jsonDoc("newName")?.ToString()

            Dim duplicated = ReformatTemplateManager.Instance.DuplicateTemplate(templateId, newName)
            If duplicated IsNot Nothing Then
                GlobalStatusStrip.ShowInfo("模板「" & duplicated.Name & "」创建成功")
                HandleGetReformatTemplates()
            Else
                GlobalStatusStrip.ShowWarning("复制模板失败")
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleDuplicateTemplate 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"复制模板失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 删除模板
    ''' </summary>
    Public Sub HandleDeleteTemplate(jsonDoc As JObject)
        Try
            Dim templateId = jsonDoc("templateId")?.ToString()

            If ReformatTemplateManager.Instance.DeleteTemplate(templateId) Then
                GlobalStatusStrip.ShowInfo("模板已删除")
                HandleGetReformatTemplates()
            Else
                GlobalStatusStrip.ShowWarning("无法删除预置模板")
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleDeleteTemplate 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"删除模板失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 打开模板编辑器
    ''' </summary>
    Public Sub HandleOpenTemplateEditor(jsonDoc As JObject)
        _invokeOnUiThread(Sub()
            Try
                Dim templateId = jsonDoc("templateId")?.ToString()
                Dim template As ReformatTemplate = Nothing

                If Not String.IsNullOrEmpty(templateId) Then
                    template = ReformatTemplateManager.Instance.GetTemplateById(templateId)
                End If

                If _showTemplateEditor(template) Then
                    Return
                End If

                Dim previewCallback = _getStylePreviewCallback()
                Dim editorForm As New ReformatTemplateEditorForm(template, previewCallback)
                If editorForm.ShowDialog() = DialogResult.OK Then
                    HandleGetReformatTemplates()
                End If
            Catch ex As Exception
                Debug.WriteLine($"HandleOpenTemplateEditor 出错: {ex.Message}")
                GlobalStatusStrip.ShowWarning($"打开模板编辑器失败: {ex.Message}")
            End Try
        End Sub)
    End Sub

    ''' <summary>
    ''' 进入模板选择模式
    ''' </summary>
    Public Async Function EnterReformatTemplateMode() As Task
        Try
            Await _executeScript("enterReformatTemplateMode();")
            Await Task.Delay(100)
            HandleGetReformatTemplates()
        Catch ex As Exception
            Debug.WriteLine($"EnterReformatTemplateMode 出错: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 退出模板选择模式
    ''' </summary>
    Public Sub ExitReformatTemplateMode()
        Try
            _executeScript("exitReformatTemplateMode();")
        Catch ex As Exception
            Debug.WriteLine($"ExitReformatTemplateMode 出错: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "AI模板编辑器"

    ''' <summary>
    ''' 处理保存AI模板
    ''' </summary>
    Public Sub HandleSaveAiTemplate(jsonDoc As JObject)
        _invokeOnUiThread(Sub()
            Try
                Dim templateJson As String = jsonDoc("templateJson")?.ToString()
                If String.IsNullOrWhiteSpace(templateJson) Then
                    templateJson = jsonDoc("template")?.ToString()
                End If
                If String.IsNullOrWhiteSpace(templateJson) Then
                    GlobalStatusStrip.ShowWarning("没有可保存的模板数据")
                    Return
                End If

                Dim template = JsonConvert.DeserializeObject(Of ReformatTemplate)(templateJson)

                If String.IsNullOrWhiteSpace(template.Id) Then
                    ReformatTemplateManager.Instance.AddTemplate(template)
                Else
                    Dim existing = ReformatTemplateManager.Instance.GetTemplateById(template.Id)
                    If existing IsNot Nothing Then
                        ReformatTemplateManager.Instance.UpdateTemplate(template)
                    Else
                        ReformatTemplateManager.Instance.AddTemplate(template)
                    End If
                End If

                GlobalStatusStrip.ShowInfo($"模板 '{template.Name}' 已保存")
                HandleGetReformatTemplates()
            Catch ex As Exception
                Debug.WriteLine($"HandleSaveAiTemplate 出错: {ex.Message}")
                GlobalStatusStrip.ShowWarning($"保存模板失败: {ex.Message}")
            End Try
        End Sub)
    End Sub

#End Region

#Region "样式规范"

    ''' <summary>
    ''' 获取排版规范列表
    ''' </summary>
    Public Sub HandleGetStyleGuides()
        Try
            Dim guides = StyleGuideManager.Instance.GetAllStyleGuides()
            Dim json = UtilsService.SerializeForScript(guides)
            _executeScript($"loadStyleGuideList({json});")
        Catch ex As Exception
            Debug.WriteLine($"HandleGetStyleGuides 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 上传规范文档
    ''' </summary>
    Public Sub HandleUploadStyleGuideDocument()
        _invokeOnUiThread(Sub()
            Try
                Dim ofd As New OpenFileDialog With {
                    .Filter = "规范文档 (*.txt;*.md;*.csv)|*.txt;*.md;*.csv|所有文件 (*.*)|*.*",
                    .Title = "选择排版规范文档"
                }

                If ofd.ShowDialog() = DialogResult.OK Then
                    Dim filePath = ofd.FileName
                    Dim detectedEncoding = DetectFileEncoding(filePath)
                    Dim content = File.ReadAllText(filePath, detectedEncoding)

                    Dim guide As New StyleGuideResource()
                    guide.Id = Guid.NewGuid().ToString()
                    guide.Name = Path.GetFileNameWithoutExtension(filePath)
                    guide.GuideContent = content
                    guide.SourceFileName = Path.GetFileName(filePath)
                    guide.SourceFileExtension = Path.GetExtension(filePath)
                    guide.FileEncoding = detectedEncoding.EncodingName
                    guide.Category = "通用"
                    guide.CreatedAt = DateTime.Now
                    guide.LastModified = DateTime.Now

                    StyleGuideManager.Instance.AddStyleGuide(guide)
                    HandleGetStyleGuides()
                    GlobalStatusStrip.ShowSuccess($"规范文档「{guide.Name}」已添加")
                End If
            Catch ex As Exception
                Debug.WriteLine($"HandleUploadStyleGuideDocument 出错: {ex.Message}")
                GlobalStatusStrip.ShowWarning($"上传规范失败: {ex.Message}")
            End Try
        End Sub)
    End Sub

    ''' <summary>
    ''' 删除规范
    ''' </summary>
    Public Sub HandleDeleteStyleGuide(jsonDoc As JObject)
        Try
            Dim guideId = jsonDoc("guideId")?.ToString()
            If StyleGuideManager.Instance.DeleteStyleGuide(guideId) Then
                HandleGetStyleGuides()
                GlobalStatusStrip.ShowSuccess("规范已删除")
            Else
                GlobalStatusStrip.ShowWarning("无法删除预置规范")
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleDeleteStyleGuide 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 更新规范内容
    ''' </summary>
    Public Sub HandleUpdateStyleGuide(jsonDoc As JObject)
        Try
            Dim guideId = jsonDoc("guideId")?.ToString()
            Dim newContent = jsonDoc("guideContent")?.ToString()
            If String.IsNullOrEmpty(guideId) Then Return

            Dim guide = StyleGuideManager.Instance.GetStyleGuideById(guideId)
            If guide Is Nothing Then Return
            If guide.IsPreset Then
                GlobalStatusStrip.ShowWarning("预置规范不可编辑")
                Return
            End If

            guide.GuideContent = newContent
            StyleGuideManager.Instance.UpdateStyleGuide(guide)
            HandleGetStyleGuides()
            GlobalStatusStrip.ShowSuccess($"规范「{guide.Name}」已保存")
        Catch ex As Exception
            Debug.WriteLine($"HandleUpdateStyleGuide 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 复制规范
    ''' </summary>
    Public Sub HandleDuplicateStyleGuide(jsonDoc As JObject)
        Try
            Dim guideId = jsonDoc("guideId")?.ToString()
            Dim newName = jsonDoc("newName")?.ToString()
            Dim duplicate = StyleGuideManager.Instance.DuplicateStyleGuide(guideId, newName)
            If duplicate IsNot Nothing Then
                HandleGetStyleGuides()
                GlobalStatusStrip.ShowSuccess($"规范「{duplicate.Name}」已创建")
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleDuplicateStyleGuide 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 导出规范
    ''' </summary>
    Public Sub HandleExportStyleGuide(jsonDoc As JObject)
        _invokeOnUiThread(Sub()
            Try
                Dim guideId = jsonDoc("guideId")?.ToString()
                Dim guide = StyleGuideManager.Instance.GetStyleGuideById(guideId)
                If guide Is Nothing Then Return

                Dim extension = If(String.IsNullOrEmpty(guide.SourceFileExtension), ".md", guide.SourceFileExtension)
                Dim sfd As New SaveFileDialog With {
                    .Filter = $"规范文件 (*{extension})|*{extension}|所有文件 (*.*)|*.*",
                    .FileName = guide.Name & extension,
                    .Title = "导出规范文档"
                }

                If sfd.ShowDialog() = DialogResult.OK Then
                    If StyleGuideManager.Instance.ExportStyleGuide(guideId, sfd.FileName) Then
                        GlobalStatusStrip.ShowSuccess($"规范已导出到: {sfd.FileName}")
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine($"HandleExportStyleGuide 出错: {ex.Message}")
            End Try
        End Sub)
    End Sub

#End Region

#Region "语义排版"

    ''' <summary>
    ''' 上传.docx模板文件并解析为SemanticStyleMapping
    ''' </summary>
    Public Sub HandleUploadDocxTemplate()
        _invokeOnUiThread(Sub()
            Try
                Dim ofd As New OpenFileDialog With {
                    .Filter = "Word模板文件 (*.docx;*.dotx)|*.docx;*.dotx|所有文件 (*.*)|*.*",
                    .Title = "选择Word模板文件"
                }

                If ofd.ShowDialog() = DialogResult.OK Then
                    _uploadDocxTemplateFromPath(ofd.FileName)
                End If
            Catch ex As Exception
                Debug.WriteLine($"HandleUploadDocxTemplate 出错: {ex.Message}")
                GlobalStatusStrip.ShowWarning($"上传模板失败: {ex.Message}")
            End Try
        End Sub)
    End Sub

    ''' <summary>
    ''' 删除docx语义映射
    ''' </summary>
    Public Sub HandleDeleteDocxMapping(jsonDoc As JObject)
        Try
            Dim mappingId = jsonDoc("mappingId")?.ToString()
            If String.IsNullOrEmpty(mappingId) Then Return

            Dim mapping = SemanticMappingManager.Instance.GetMappingById(mappingId)
            If mapping IsNot Nothing Then
                If Not String.IsNullOrEmpty(mapping.SourceFilePath) AndAlso IO.File.Exists(mapping.SourceFilePath) Then
                    Try
                        IO.File.Delete(mapping.SourceFilePath)
                    Catch ex As Exception
                        Debug.WriteLine($"删除模板文件失败: {ex.Message}")
                    End Try
                End If
                SemanticMappingManager.Instance.DeleteMapping(mappingId)
            End If

            HandleGetReformatTemplates()
            GlobalStatusStrip.ShowInfo("已删除文档映射")
        Catch ex As Exception
            Debug.WriteLine($"HandleDeleteDocxMapping 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"删除映射失败: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "排版重试"

    ''' <summary>
    ''' 获取重试次数
    ''' </summary>
    Public Function GetRetryCount(uuid As String) As Integer
        Dim count As Integer = 0
        _reformatRetryCount.TryGetValue(uuid, count)
        Return count
    End Function

    ''' <summary>
    ''' 增加重试次数并返回新值
    ''' </summary>
    Public Function IncrementRetryCount(uuid As String) As Integer
        Dim count = GetRetryCount(uuid)
        _reformatRetryCount(uuid) = count + 1
        Return count + 1
    End Function

#End Region

#Region "工具方法"

    ''' <summary>
    ''' 自动检测文件编码（支持BOM检测和GBK回退）
    ''' </summary>
    Private Function DetectFileEncoding(filePath As String) As System.Text.Encoding
        Try
            Dim bytes = File.ReadAllBytes(filePath)
            If bytes.Length = 0 Then Return System.Text.Encoding.UTF8

            If bytes.Length >= 3 AndAlso bytes(0) = &HEF AndAlso bytes(1) = &HBB AndAlso bytes(2) = &HBF Then
                Return System.Text.Encoding.UTF8
            End If
            If bytes.Length >= 2 AndAlso bytes(0) = &HFF AndAlso bytes(1) = &HFE Then
                Return System.Text.Encoding.Unicode
            End If
            If bytes.Length >= 2 AndAlso bytes(0) = &HFE AndAlso bytes(1) = &HFF Then
                Return System.Text.Encoding.BigEndianUnicode
            End If

            Try
                Dim utf8 As New System.Text.UTF8Encoding(False, True)
                utf8.GetString(bytes)
                Return System.Text.Encoding.UTF8
            Catch ex As System.Text.DecoderFallbackException
            End Try

            Try
                Return System.Text.Encoding.GetEncoding("GBK")
            Catch
                Return System.Text.Encoding.Default
            End Try
        Catch ex As Exception
            Debug.WriteLine($"编码检测失败: {ex.Message}")
            Return System.Text.Encoding.UTF8
        End Try
    End Function

#End Region

#Region "智能排版（v2）"

    ''' <summary>
    ''' 一键速排：分析文档 → 推荐标准 → 生成预览
    ''' </summary>
    Public Async Function QuickReformatAsync(
        paragraphs As List(Of String),
        wordParagraphs As List(Of Object)) As Task(Of ReformatPreviewPlan)

        Try
            GlobalStatusStrip.ShowInfo("正在分析文档...")
            Dim plan = Await FormattingOrchestrator.QuickReformatAsync(paragraphs, wordParagraphs)

            If plan.Changes.Count = 0 Then
                GlobalStatusStrip.ShowWarning("未找到匹配的排版方案，请尝试使用模板")
            Else
                GlobalStatusStrip.ShowSuccess($"分析完成，发现{plan.TotalChanges}处可优化项")
            End If

            Return plan
        Catch ex As Exception
            Debug.WriteLine($"QuickReformatAsync 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"排版分析失败: {ex.Message}")
            Return New ReformatPreviewPlan()
        End Try
    End Function

    ''' <summary>
    ''' 对话式排版：解析用户自然语言指令
    ''' </summary>
    Public Async Function ChatReformatAsync(
        userMessage As String,
        paragraphs As List(Of String),
        wordParagraphs As List(Of Object)) As Task(Of ReformatPreviewPlan)

        Try
            Dim responseUuid As String = ""
            Dim handled = Await ChatFormatterAgent.HandleFormattingMessage(
                userMessage, paragraphs, wordParagraphs, responseUuid)

            If handled Then
                Return ChatFormatterAgent.Orchestrator.RefinementContext.LastPlan
            End If

            ' 非排版消息，执行默认分析
            Return Await QuickReformatAsync(paragraphs, wordParagraphs)
        Catch ex As Exception
            Debug.WriteLine($"ChatReformatAsync 出错: {ex.Message}")
            Return New ReformatPreviewPlan()
        End Try
    End Function

#End Region

End Class
