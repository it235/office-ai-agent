Imports System.IO
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 排版服务：排版模板、样式规范、AI模板编辑器、语义排版（非Overridable方法）
''' </summary>
Public Class ReformatService

    Private ReadOnly _executeScript As Func(Of String, Task)
    Private ReadOnly _escapeJs As Func(Of String, String)
    Private ReadOnly _invokeOnUiThread As Action(Of Action)
    Private ReadOnly _showTemplateEditor As Func(Of ReformatTemplate, Boolean)
    Private ReadOnly _getStylePreviewCallback As Func(Of PreviewStyleCallback)
    Private ReadOnly _uploadDocxTemplateFromPath As Action(Of String)

    ' 排版重试上下文
    Private _reformatRetryContext As New Dictionary(Of String, Tuple(Of String, String))()
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
                _chatFormatterAgent = New ChatFormatterAgent(_executeScript, _escapeJs)
            End If
            Return _chatFormatterAgent
        End Get
    End Property

    Public Sub New(
        executeScript As Func(Of String, Task),
        escapeJs As Func(Of String, String),
        invokeOnUiThread As Action(Of Action),
        showTemplateEditor As Func(Of ReformatTemplate, Boolean),
        getStylePreviewCallback As Func(Of PreviewStyleCallback),
        uploadDocxTemplateFromPath As Action(Of String))

        _executeScript = executeScript
        _escapeJs = escapeJs
        _invokeOnUiThread = invokeOnUiThread
        _showTemplateEditor = showTemplateEditor
        _getStylePreviewCallback = getStylePreviewCallback
        _uploadDocxTemplateFromPath = uploadDocxTemplateFromPath
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
                allItems.Add(t)
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

            Dim json = JsonConvert.SerializeObject(allItems, Formatting.None)
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
    ''' 进入AI模板编辑模式
    ''' </summary>
    Public Sub EnterAiTemplateEditorMode(Optional template As ReformatTemplate = Nothing)
        Try
            Dim templateJson As String = ""
            If template IsNot Nothing Then
                templateJson = JsonConvert.SerializeObject(template)
            End If
            _executeScript($"enterAiTemplateEditor('{_escapeJs(templateJson)}');")
        Catch ex As Exception
            Debug.WriteLine($"EnterAiTemplateEditorMode 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理开始AI模板创建对话
    ''' </summary>
    Public Sub HandleStartAiTemplateChat(jsonDoc As JObject)
        _invokeOnUiThread(Sub()
            Try
                Dim mode As String = jsonDoc("mode")?.ToString()
                Dim promptMessage As String

                If mode = "fromSelection" Then
                    promptMessage = "请帮我根据当前文档的排版样式创建一个ReformatTemplate模板。" & vbCrLf &
                                   "请分析文档中的标题、正文、段落格式等，生成一个完整的JSON格式模板。" & vbCrLf &
                                   "模板必须包含Name、Layout、BodyStyles、PageSettings字段。"
                Else
                    promptMessage = "我想创建一个文档排版模板（ReformatTemplate）。" & vbCrLf &
                                   "请问你想创建什么类型的文档模板？（如：公文、论文、报告、简历等）" & vbCrLf &
                                   "请告诉我模板的用途，我会帮你生成一个完整的JSON格式模板。"
                End If

                Dim escapedPrompt = _escapeJs(promptMessage)
                _executeScript($"document.getElementById('message-input').value = '{escapedPrompt}'; document.getElementById('message-input').focus();")
                GlobalStatusStrip.ShowInfo("请在聊天框中描述您需要的模板类型，AI将为您生成模板")
            Catch ex As Exception
                Debug.WriteLine($"HandleStartAiTemplateChat 出错: {ex.Message}")
                GlobalStatusStrip.ShowWarning($"启动AI模板对话失败: {ex.Message}")
            End Try
        End Sub)
    End Sub

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
            Dim json = JsonConvert.SerializeObject(guides, Formatting.None)
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
    ''' 保存排版请求上下文，用于重试
    ''' </summary>
    Public Sub SaveReformatContext(uuid As String, systemPrompt As String, userMessage As String)
        _reformatRetryContext(uuid) = Tuple.Create(systemPrompt, userMessage)
        _reformatRetryCount(uuid) = 0
    End Sub

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

    ''' <summary>
    ''' 获取文档分析结果（从JS请求触发）
    ''' 从Word获取段落文本 → DocumentAnalyzer分析 → 推送结果到JS
    ''' </summary>
    Public Sub HandleAnalyzeDocument(jsonDoc As JObject)
        Try
            Dim paragraphsJson = jsonDoc("paragraphs")?.ToString()
            If String.IsNullOrEmpty(paragraphsJson) Then
                GlobalStatusStrip.ShowWarning("未获取到文档段落数据")
                Return
            End If

            Dim paragraphs = JsonConvert.DeserializeObject(Of List(Of String))(paragraphsJson)
            If paragraphs Is Nothing OrElse paragraphs.Count = 0 Then
                GlobalStatusStrip.ShowWarning("文档段落为空")
                Return
            End If

            Dim analyzer As New DocumentAnalyzer()
            Dim result = analyzer.Analyze(paragraphs)

            Dim output As New JObject()
            output("docType") = result.DocumentType.ToString()
            output("docTypeName") = GetDocumentTypeNameChinese(result.DocumentType)
            output("confidence") = result.Confidence
            output("paragraphCount") = result.ParagraphCount
            output("hasToc") = result.HasTableOfContents
            output("analysisTimeMs") = result.AnalysisTimeMs

            ' 格式问题
            Dim problemsArray As New JArray()
            For Each p In result.FormattingProblems
                Dim item As New JObject()
                item("description") = p.Description
                item("severity") = p.Severity.ToString()
                item("category") = p.Category
                item("suggestedFix") = p.SuggestedFix
                problemsArray.Add(item)
            Next
            output("problems") = problemsArray

            ' 标题结构
            Dim headingsArray As New JArray()
            If result.DocStructure IsNot Nothing Then
                For Each h In result.DocStructure.Headings
                    Dim item As New JObject()
                    item("level") = h.Level
                    item("text") = h.Text
                    item("paragraphIndex") = h.ParagraphIndex
                    headingsArray.Add(item)
                Next
            End If
            output("headings") = headingsArray

            _executeScript($"showDocumentAnalysis({output.ToString(Formatting.None)});")
            GlobalStatusStrip.ShowInfo($"文档分析完成: {GetDocumentTypeNameChinese(result.DocumentType)}({Math.Round(result.Confidence * 100)}%)")

        Catch ex As Exception
            Debug.WriteLine($"HandleAnalyzeDocument 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"文档分析失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 生成排版预览（不实际应用）
    ''' 将排版方案以卡片形式推送到JS前端展示
    ''' </summary>
    Public Async Function HandlePreviewReformat(jsonDoc As JObject) As Task
        Try
            Dim paragraphsJson = jsonDoc("paragraphs")?.ToString()

            If String.IsNullOrEmpty(paragraphsJson) Then
                GlobalStatusStrip.ShowWarning("未获取到段落数据")
                Return
            End If

            Dim paragraphs = JsonConvert.DeserializeObject(Of List(Of String))(paragraphsJson)

            GlobalStatusStrip.ShowInfo("正在分析文档...")
            Dim plan = Await QuickReformatAsync(paragraphs, New List(Of Object)())

            Dim html = ChatFormatterAgent.GenerateFormattingCardHtml(plan)
            Dim escapedHtml = _escapeJs(html)

            Await _executeScript($"showFormattingPreview('{escapedHtml}');")
            GlobalStatusStrip.ShowInfo("排版预览已生成")

        Catch ex As Exception
            Debug.WriteLine($"HandlePreviewReformat 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"生成排版预览失败: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 应用排版（从预览确认后执行）
    ''' 将语义排版映射应用到Word文档
    ''' </summary>
    Public Async Function HandleApplyReformat(jsonDoc As JObject) As Task
        Try
            Dim mappingJson = jsonDoc("mapping")?.ToString()
            Dim taggedParagraphsJson = jsonDoc("taggedParagraphs")?.ToString()
            Dim wordParagraphsJson = jsonDoc("wordParagraphs")?.ToString()

            If String.IsNullOrEmpty(mappingJson) Then
                ' 使用当前上下文中最后的排版方案
                Dim context = FormattingOrchestrator.RefinementContext
                If context.LastPlan IsNot Nothing AndAlso context.LastPlan.SemanticMapping IsNot Nothing Then
                    mappingJson = JsonConvert.SerializeObject(context.LastPlan.SemanticMapping)
                Else
                    GlobalStatusStrip.ShowWarning("没有可应用的排版方案")
                    Return
                End If
            End If

            Dim mapping = JsonConvert.DeserializeObject(Of SemanticStyleMapping)(mappingJson)
            If mapping Is Nothing Then
                GlobalStatusStrip.ShowWarning("排版映射数据无效")
                Return
            End If

            ' 如果没有传入标注结果，构建默认标注（全文为正文）
            Dim taggedParagraphs As List(Of TaggedParagraph)
            If Not String.IsNullOrEmpty(taggedParagraphsJson) Then
                taggedParagraphs = JsonConvert.DeserializeObject(Of List(Of TaggedParagraph))(taggedParagraphsJson)
            Else
                ' 从JSON中解析wordParagraphs数量
                Dim paraCountObj = jsonDoc("paraCount")?.ToObject(Of Integer)()
                Dim paraCount = If(paraCountObj, 0)
                taggedParagraphs = New List(Of TaggedParagraph)()
                For i = 0 To paraCount - 1
                    taggedParagraphs.Add(New TaggedParagraph(i, "body.normal"))
                Next
            End If

            ' 构建段落类型列表：优先从JSON获取，其次从编排器上下文获取，最后从Word推断
            Dim paragraphTypes As List(Of String) = Nothing

            ' 1. 尝试从JSON payload获取
            Dim paragraphTypesJson = jsonDoc("paragraphTypes")?.ToString()
            If Not String.IsNullOrEmpty(paragraphTypesJson) Then
                paragraphTypes = JsonConvert.DeserializeObject(Of List(Of String))(paragraphTypesJson)
            End If

            ' 2. 尝试从编排器上下文获取
            If paragraphTypes Is Nothing Then
                Dim context = FormattingOrchestrator.RefinementContext
                If context.LastPlan IsNot Nothing AndAlso context.LastPlan.ParagraphTypes IsNot Nothing Then
                    paragraphTypes = context.LastPlan.ParagraphTypes
                End If
            End If

            ' 3. 从Word段落对象推断（在UI线程中执行）
            If paragraphTypes Is Nothing Then
                paragraphTypes = New List(Of String)()
                For rt = 0 To taggedParagraphs.Count - 1
                    paragraphTypes.Add("text")
                Next
            End If

            ' 执行渲染（在UI线程中操作Word对象）
            _invokeOnUiThread(Sub()
                Try
                    ' 获取Word Application对象（通过反射访问）
                    Dim wordApp = GetWordApplication()
                    If wordApp Is Nothing Then
                        GlobalStatusStrip.ShowWarning("无法访问Word应用程序")
                        Return
                    End If

                    ' 获取Word段落对象列表
                    Dim wordParagraphs As New List(Of Object)()
                    Dim doc = wordApp.ActiveDocument
                    For i = 1 To doc.Paragraphs.Count
                        wordParagraphs.Add(doc.Paragraphs.Item(i))
                    Next

                    Dim result = SemanticRenderingEngine.ApplySemanticFormatting(
                        taggedParagraphs, mapping, wordParagraphs, paragraphTypes, wordApp)

                    FormattingOrchestrator.RefinementContext.IsApplied = True

                    Dim output As New JObject()
                    output("appliedCount") = result.AppliedCount
                    output("skippedCount") = result.SkippedCount
                    _executeScript($"onReformatApplied({output.ToString(Formatting.None)});")

                    GlobalStatusStrip.ShowSuccess($"排版应用完成: {result.AppliedCount}段已修改")
                Catch ex As Exception
                    Debug.WriteLine($"应用排版失败: {ex.Message}")
                    GlobalStatusStrip.ShowWarning($"应用排版失败: {ex.Message}")
                End Try
            End Sub)

        Catch ex As Exception
            Debug.WriteLine($"HandleApplyReformat 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"应用排版失败: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 对话微调：在当前排版方案基础上做增量修改
    ''' </summary>
    Public Async Function HandleRefineReformat(jsonDoc As JObject) As Task
        Try
            Dim userMessage = jsonDoc("message")?.ToString()
            Dim paragraphsJson = jsonDoc("paragraphs")?.ToString()

            If String.IsNullOrEmpty(userMessage) Then
                GlobalStatusStrip.ShowWarning("未指定微调指令")
                Return
            End If

            ' 确保有进行中的排版上下文
            If Not FormattingOrchestrator.HasActiveContext() Then
                GlobalStatusStrip.ShowWarning("没有可微调的排版方案，请先执行排版分析")
                Return
            End If

            Dim paragraphs = If(Not String.IsNullOrEmpty(paragraphsJson),
                JsonConvert.DeserializeObject(Of List(Of String))(paragraphsJson),
                New List(Of String)())

            ' ChatReformatAsync internally parses intent and validates
            Await FormattingOrchestrator.ChatReformatAsync(userMessage, paragraphs, New List(Of Object)())
            Dim refinedPlan = FormattingOrchestrator.RefinementContext.LastPlan

            If refinedPlan IsNot Nothing Then
                ' 推送预览卡片到前端
                Dim json = refinedPlan.ToPreviewJson().ToString(Newtonsoft.Json.Formatting.None)
                _executeScript($"showFormattingPreview({json});")
                GlobalStatusStrip.ShowSuccess("排版微调已应用，请预览确认")
            End If

        Catch ex As Exception
            Debug.WriteLine($"HandleRefineReformat 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"排版微调失败: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 格式克隆：对比范文与当前文档，提取格式并应用
    ''' </summary>
    Public Async Function HandleMirrorFormat(jsonDoc As JObject) As Task
        Try
            Dim referenceDocPath = jsonDoc("referencePath")?.ToString()

            If String.IsNullOrEmpty(referenceDocPath) Then
                ' 打开文件对话框选择范文
                _invokeOnUiThread(Sub()
                    Dim ofd As New OpenFileDialog With {
                        .Filter = "Word文档 (*.docx;*.doc)|*.docx;*.doc|所有文件 (*.*)|*.*",
                        .Title = "选择范文文档"
                    }
                    If ofd.ShowDialog() = DialogResult.OK Then
                        referenceDocPath = ofd.FileName
                        ProcessMirrorFormatInternal(referenceDocPath)
                    End If
                End Sub)
            Else
                ProcessMirrorFormatInternal(referenceDocPath)
            End If

        Catch ex As Exception
            Debug.WriteLine($"HandleMirrorFormat 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"格式克隆失败: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 内部处理格式克隆逻辑
    ''' </summary>
    Private Async Sub ProcessMirrorFormatInternal(referenceDocPath As String)
        Try
            GlobalStatusStrip.ShowInfo("正在分析范文格式...")

            ' 获取Word Application
            Dim wordApp = GetWordApplication()
            If wordApp Is Nothing Then
                GlobalStatusStrip.ShowWarning("无法访问Word应用程序")
                Return
            End If

            ' 打开范文文档（在后台打开，不显示）
            Dim refDoc As Object = Nothing
            Try
                refDoc = wordApp.Documents.Open(referenceDocPath, ReadOnly:=True, Visible:=False)
            Catch ex As Exception
                GlobalStatusStrip.ShowWarning($"无法打开范文: {ex.Message}")
                Return
            End Try

            Try
                ' 提取范文格式
                Dim extractedFormats = FormatMirrorService.ExtractFormattingFromDocument(wordApp, False)
                If extractedFormats Is Nothing OrElse extractedFormats.Count = 0 Then
                    GlobalStatusStrip.ShowWarning("未能从范文中提取格式信息")
                    Return
                End If

                ' 关闭范文
                refDoc.Close(SaveChanges:=False)

                ' 构建AI克隆提示词
                Dim clonePrompt = FormatMirrorService.BuildClonePrompt(extractedFormats)

                ' 通过JS传递到AI处理
                Dim promptJson As New JObject()
                promptJson("prompt") = clonePrompt
                promptJson("extractedCount") = extractedFormats.Count

                _executeScript($"onMirrorFormatReady({promptJson.ToString(Formatting.None)});")
                GlobalStatusStrip.ShowSuccess($"已从范文提取{extractedFormats.Count}种格式规则，请确认是否应用")

            Catch ex As Exception
                ' 确保范文关闭
                Try
                    refDoc?.Close(SaveChanges:=False)
                Catch
                End Try
                Debug.WriteLine($"格式克隆处理失败: {ex.Message}")
                GlobalStatusStrip.ShowWarning($"格式克隆失败: {ex.Message}")
            End Try

        Catch ex As Exception
            Debug.WriteLine($"ProcessMirrorFormatInternal 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"格式克隆失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 获取Word Application对象（通过反射，避免ShareRibbon直接依赖Word Interop）
    ''' </summary>
    Private Shared Function GetWordApplication() As Object
        Try
            Dim wordApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application")
            Return wordApp
        Catch ex As Exception
            Debug.WriteLine($"获取Word Application失败: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 获取文档类型中文名称
    ''' </summary>
    Private Shared Function GetDocumentTypeNameChinese(docType As DocumentType) As String
        Select Case docType
            Case DocumentType.OfficialDocument : Return "行政公文"
            Case DocumentType.AcademicPaper : Return "学术论文"
            Case DocumentType.BusinessReport : Return "商业报告"
            Case DocumentType.Contract : Return "合同协议"
            Case DocumentType.Resume : Return "个人简历"
            Case DocumentType.GeneralDocument : Return "通用文档"
            Case Else : Return "未知类型"
        End Select
    End Function

#End Region

End Class
