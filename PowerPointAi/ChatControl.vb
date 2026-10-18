Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Mime
Imports System.Reflection.Emit
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Windows.Forms
Imports System.Windows.Forms.ListBox
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports PowerPointAi.Extensions
Imports ShareRibbon.Extensions
Imports ShareRibbon
Public Class ChatControl
    Inherits BaseChatControl



    ' 排版上下文：存储待格式化的形状和类型信息
    Private _reformatShapes As List(Of Object) = Nothing
    Private _reformatTypes As List(Of String) = Nothing

    ' 排版撤销快照（PPT不支持UndoRecord，使用自定义快照）
    Private _reformatUndoSnapshots As List(Of ShapeFormatSnapshot) = Nothing

    ''' <summary>
    ''' 形状格式快照 - 用于PPT排版撤销
    ''' </summary>
    Private Class ShapeFormatSnapshot
        Public ShapeIndex As Integer
        Public FontNameFarEast As String = ""
        Public FontName As String = ""
        Public FontSize As Single = 0
        Public Bold As Microsoft.Office.Core.MsoTriState = Microsoft.Office.Core.MsoTriState.msoFalse
        Public Alignment As Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft

        Public Shared Function Capture(shp As Microsoft.Office.Interop.PowerPoint.Shape, index As Integer) As ShapeFormatSnapshot
            Dim snap As New ShapeFormatSnapshot()
            snap.ShapeIndex = index
            Try
                If shp.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    Dim tr = shp.TextFrame.TextRange
                    snap.FontNameFarEast = If(tr.Font.NameFarEast IsNot Nothing, tr.Font.NameFarEast, "")
                    snap.FontName = If(tr.Font.Name IsNot Nothing, tr.Font.Name, "")
                    snap.FontSize = tr.Font.Size
                    snap.Bold = tr.Font.Bold
                    snap.Alignment = tr.ParagraphFormat.Alignment
                End If
            Catch ex As Exception
                Debug.WriteLine($"捕获形状{index}快照失败: {ex.Message}")
            End Try
            Return snap
        End Function

        Public Sub Restore(shp As Microsoft.Office.Interop.PowerPoint.Shape)
            Try
                If shp.HasTextFrame <> Microsoft.Office.Core.MsoTriState.msoTrue Then Return
                Dim tr = shp.TextFrame.TextRange
                If Not String.IsNullOrEmpty(FontNameFarEast) Then tr.Font.NameFarEast = FontNameFarEast
                If Not String.IsNullOrEmpty(FontName) Then tr.Font.Name = FontName
                If FontSize > 0 Then tr.Font.Size = FontSize
                tr.Font.Bold = Bold
                tr.ParagraphFormat.Alignment = Alignment
            Catch ex As Exception
                Debug.WriteLine($"恢复形状{ShapeIndex}快照失败: {ex.Message}")
            End Try
        End Sub
    End Class

    ''' <summary>
    ''' 设置排版上下文，用于规则匹配后应用格式
    ''' </summary>
    Public Sub SetReformatContext(shapes As List(Of Object), types As List(Of String))
        ' 新排版开始时，清空旧的撤销快照（防止不匹配）
        _reformatUndoSnapshots = Nothing
        _reformatShapes = shapes
        _reformatTypes = types
    End Sub

    ''' <summary>
    ''' 使用模板进行排版（覆盖基类方法）
    ''' </summary>
    Protected Overrides Async Sub ApplyReformatWithTemplate(template As ReformatTemplate)
        Try
            Dim pptApp = Globals.ThisAddIn.Application
            Dim activeWindow = pptApp.ActiveWindow
            Dim selection = activeWindow.Selection

            ' 收集所有待排版的形状
            Dim selectedShapes As New List(Of Microsoft.Office.Interop.PowerPoint.Shape)()
            Dim shapeTypes As New List(Of String)()

            Try
                If selection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes Then
                    ' 选中了形状
                    For Each shp As Microsoft.Office.Interop.PowerPoint.Shape In selection.ShapeRange
                        If shp.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue AndAlso
                           shp.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            selectedShapes.Add(shp)
                            shapeTypes.Add(GetShapeTypeName(shp))
                        End If
                    Next
                ElseIf selection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionSlides Then
                    ' 选中了幻灯片，获取所有文本形状
                    For Each slide As Microsoft.Office.Interop.PowerPoint.Slide In selection.SlideRange
                        For Each shp As Microsoft.Office.Interop.PowerPoint.Shape In slide.Shapes
                            If shp.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue AndAlso
                               shp.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                selectedShapes.Add(shp)
                                shapeTypes.Add(GetShapeTypeName(shp))
                            End If
                        Next
                    Next
                End If
            Catch ex As Exception
                Debug.WriteLine("获取选中内容失败: " & ex.Message)
            End Try

            If selectedShapes.Count = 0 Then
                GlobalStatusStrip.ShowWarning("请先选中需要排版的幻灯片或文本框。")
                Return
            End If

            ' 统计形状类型
            Dim titleCount = shapeTypes.Where(Function(t) t.Contains("标题")).Count()
            Dim bodyCount = shapeTypes.Where(Function(t) t = "正文" OrElse t = "文本框").Count()

            ' 采样策略：只取代表性样本（最多5个）
            Dim sampleBlocks As New Newtonsoft.Json.Linq.JArray()
            Dim sampleIndices As New List(Of Integer)()
            Dim totalCount = selectedShapes.Count

            If totalCount <= 5 Then
                For i As Integer = 0 To totalCount - 1
                    sampleIndices.Add(i)
                Next
            Else
                ' 采样：首2、中1、尾2
                sampleIndices.Add(0)
                sampleIndices.Add(1)
                sampleIndices.Add(CInt(Math.Floor(totalCount / 2)))
                sampleIndices.Add(totalCount - 2)
                sampleIndices.Add(totalCount - 1)
            End If

            For Each idx In sampleIndices
                Dim shp = selectedShapes(idx)
                Dim textContent = ""
                Try
                    textContent = shp.TextFrame.TextRange.Text
                    ' 头尾采样：截断过长文本
                    If textContent.Length > 80 Then
                        textContent = textContent.Substring(0, 40) & "..." & textContent.Substring(textContent.Length - 30)
                    End If
                Catch
                End Try

                Dim sampleObj As New Newtonsoft.Json.Linq.JObject()
                sampleObj("sampleIndex") = idx
                sampleObj("text") = textContent
                sampleObj("currentType") = shapeTypes(idx)
                sampleBlocks.Add(sampleObj)
            Next

            ' 显示排版模式吸顶提示
            Await ExecuteJavaScriptAsyncJS("showReformatModeIndicator();")

            ' 构建带模板的系统提示
            Dim systemPrompt As New System.Text.StringBuilder()
            systemPrompt.AppendLine("你是PowerPoint排版助手。用户选择了「" & template.Name & "」模板进行排版。")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("【模板配置】")
            systemPrompt.AppendLine($"模板名称：{template.Name}")
            systemPrompt.AppendLine($"模板分类：{template.Category}")
            systemPrompt.AppendLine($"模板描述：{template.Description}")
            systemPrompt.AppendLine()

            ' 版式配置
            If template.Layout IsNot Nothing AndAlso template.Layout.Elements IsNot Nothing AndAlso template.Layout.Elements.Count > 0 Then
                systemPrompt.AppendLine("版式骨架元素：")
                For Each el In template.Layout.Elements
                    systemPrompt.AppendLine($"  - {el.Name}: {el.Font?.FontNameCN} {el.Font?.FontSize}pt, {el.Paragraph?.Alignment}")
                Next
                systemPrompt.AppendLine()
            End If

            ' 正文样式
            If template.BodyStyles IsNot Nothing AndAlso template.BodyStyles.Count > 0 Then
                systemPrompt.AppendLine("正文样式规则：")
                For Each style In template.BodyStyles
                    systemPrompt.AppendLine($"  - {style.RuleName}: {style.Font?.FontNameCN} {style.Font?.FontSize}pt")
                Next
                systemPrompt.AppendLine()
            End If

            ' AI说明
            If Not String.IsNullOrEmpty(template.AiGuidance) Then
                systemPrompt.AppendLine("【模板说明】")
                systemPrompt.AppendLine(template.AiGuidance)
                systemPrompt.AppendLine()
            End If

            systemPrompt.AppendLine("【文档信息】")
            systemPrompt.AppendLine($"演示文稿共有{totalCount}个文本框（{titleCount}个标题，{bodyCount}个正文/文本框）。")
            systemPrompt.AppendLine($"我发送了{sampleIndices.Count}个代表性样本给你。")
            systemPrompt.AppendLine()

            systemPrompt.AppendLine("【任务要求】")
            systemPrompt.AppendLine("请根据模板配置和文本框样本，返回具体的排版规则JSON。格式如下：")
            systemPrompt.AppendLine("```json")
            systemPrompt.AppendLine("{")
            systemPrompt.AppendLine("  ""rules"": [{""type"": ""title"", ""matchCondition"": ""..."", ""formatting"": {""fontNameCN"": ""黑体"", ""fontSize"": 36, ""bold"": true, ""alignment"": ""center""}}],")
            systemPrompt.AppendLine("  ""sampleClassification"": [{""sampleIndex"": 0, ""appliedRule"": ""title""}],")
            systemPrompt.AppendLine("  ""summary"": ""排版策略说明""")
            systemPrompt.AppendLine("}")
            systemPrompt.AppendLine("```")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("formatting字段说明：fontNameCN(中文字体), fontNameEN(英文字体), fontSize(字号pt), bold(加粗), alignment(对齐left/center/right)")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("以下是采样的文本框样本：")
            systemPrompt.AppendLine(sampleBlocks.ToString(Newtonsoft.Json.Formatting.Indented))

            ' 保存上下文用于后续应用
            SetReformatContext(selectedShapes.Cast(Of Object).ToList(), shapeTypes)

            ' 发送请求
            Await Send("请使用「" & template.Name & "」模板对选中内容进行排版。", systemPrompt.ToString(), False, "reformat")

            GlobalStatusStrip.ShowInfo("正在使用「" & template.Name & "」模板排版...")

        Catch ex As Exception
            Debug.WriteLine($"ApplyReformatWithTemplate 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"排版失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 获取形状类型名称（用于模板排版）
    ''' </summary>
    Private Function GetShapeTypeName(shp As Microsoft.Office.Interop.PowerPoint.Shape) As String
        Try
            If shp.PlaceholderFormat IsNot Nothing Then
                Select Case shp.PlaceholderFormat.Type
                    Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle,
                         Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle
                        Return "标题"
                    Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle
                        Return "副标题"
                    Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderBody
                        Return "正文"
                End Select
            End If
        Catch
        End Try
        Return "文本框"
    End Function

    ''' <summary>
    ''' 获取当前 Office 应用程序名称
    ''' </summary>
    Protected Overrides Function GetOfficeApplicationName() As String
        Return "PowerPoint"
    End Function


    Public Sub New()
        ' 此调用是设计师所必需的。
        InitializeComponent()

        ' 确保WebView2控件可以正常交互
        ChatBrowser.BringToFront()

        '加入底部告警栏
        Me.Controls.Add(GlobalStatusStrip.StatusStrip)

        ' 订阅Word的SelectionChange 事件
        ' 帮我补全word选择的内容事件
        AddHandler Globals.ThisAddIn.Application.WindowSelectionChange, AddressOf GetSelectionContent
    End Sub

    '获取选中的内容
    Protected Overrides Sub GetSelectionContent(target As Object)
        Try
            If Not Me.Visible OrElse Not selectedCellChecked Then
                Return
            End If

            ' 转换为 PowerPoint.Selection 对象
            Dim selection = Globals.ThisAddIn.Application.ActiveWindow.Selection
            If selection Is Nothing Then
                Return
            End If

            ' 获取选中内容的详细信息
            Dim content As String = String.Empty

            ' 根据选择类型处理内容
            If selection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes Then
                ' 处理形状选择
                Dim shapeRange = selection.ShapeRange
                If shapeRange.Count > 0 Then
                    ' 检查是否是表格
                    If shapeRange(1).HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                        ' 处理表格
                        Dim table = shapeRange(1).Table
                        Dim sb As New StringBuilder()
                        For row As Integer = 1 To table.Rows.Count
                            For col As Integer = 1 To table.Columns.Count
                                sb.Append(table.Cell(row, col).Shape.TextFrame.TextRange.Text.Trim())
                                If col < table.Columns.Count Then sb.Append(vbTab)
                            Next
                            sb.AppendLine()
                        Next
                        content = sb.ToString()
                    Else
                        ' 处理普通形状
                        content = "[已选中 " & shapeRange.Count & " 个形状]"
                        For i = 1 To shapeRange.Count
                            If shapeRange(i).HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                content &= vbCrLf & shapeRange(i).TextFrame.TextRange.Text
                            End If
                        Next
                    End If
                End If

            ElseIf selection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText Then
                ' 处理文本选择
                content = selection.TextRange.Text

            ElseIf selection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionSlides Then
                ' 处理幻灯片选择
                content = "[已选中 " & selection.SlideRange.Count & " 张幻灯片]"
            End If

            If Not String.IsNullOrEmpty(content) Then
                ' 添加到选中内容列表
                AddSelectedContentItem(
                "PowerPoint幻灯片",  ' 使用文档名称作为标识
                content.Substring(0, Math.Min(content.Length, 50)) & If(content.Length > 50, "...", "")
            )
            Else
                ' 选中没有内容，清除相同 sheetName 的引用
                ClearSelectedContentBySheetName("PowerPoint幻灯片")
            End If

        Catch ex As Exception
            Debug.WriteLine($"获取PowerPoint选中内容时出错: {ex.Message}")
        End Try
    End Sub

    ' 初始化时注入基础 HTML 结构
    Private Async Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 初始化 WebView2
        Await InitializeWebView2()
        InitializeWebView2Script()
    End Sub


    Protected Overrides Function GetVBProject() As VBProject
        Try
            Dim project = Globals.ThisAddIn.Application.VBE.ActiveVBProject
            Return project
        Catch ex As Runtime.InteropServices.COMException
            VBAxceptionHandle(ex)
            Return Nothing
        End Try
    End Function

    Protected Overrides Function RunCode(code As String) As Object
        Try
            Globals.ThisAddIn.Application.Run(code)
            Return True
        Catch ex As Runtime.InteropServices.COMException
            VBAxceptionHandle(ex)
            Return False
        Catch ex As Exception
            MessageBox.Show("执行代码时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' 执行前预览代码
    Protected Overrides Function RunCodePreview(vbaCode As String, preview As Boolean) As Boolean
        Return True
    End Function

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("PowerPoint", OfficeApplicationType.PowerPoint)
    End Function

    ' 返回Office应用类型
    Protected Overrides Function GetOfficeAppType() As String
        Return "PowerPoint"
    End Function

    ' 提供PowerPoint应用程序对象
    Protected Overrides Function GetOfficeApplicationObject() As Object
        Return Globals.ThisAddIn.Application
    End Function

    Protected Overrides Sub SendChatMessage(message As String)
        ' 这里可以实现word的特殊逻辑
        Send(message, "", True, "")
    End Sub

    ''' <summary>
    ''' 使用意图识别结果发送聊天消息（重写基类方法）
    ''' </summary>
    Protected Overrides Sub SendChatMessageWithIntent(message As String, intent As IntentResult)
        If intent IsNot Nothing AndAlso intent.Confidence > 0.2 Then
            Dim optimizedPrompt = IntentService.GetOptimizedSystemPrompt(intent)
            Debug.WriteLine($"PPT使用意图优化提示词: {intent.IntentType}, 置信度: {intent.Confidence:F2}")

            Task.Run(Async Function()
                         Await Send(message, optimizedPrompt, True, "")
                     End Function)
        Else
            ' 回退到普通发送
            SendChatMessage(message)
        End If
    End Sub

    Protected Overrides Function ParseFile(filePath As String) As FileContentResult
        Try
            ' 复用当前 PPT 进程，WithWindow=msoFalse 静默打开，避免创建第二个 PPT 实例
            Dim pptApp = Globals.ThisAddIn.Application

            Dim presentation As Microsoft.Office.Interop.PowerPoint.Presentation = Nothing
            Try
                presentation = pptApp.Presentations.Open(filePath,
                    ReadOnly:=Microsoft.Office.Core.MsoTriState.msoTrue,
                    Untitled:=Microsoft.Office.Core.MsoTriState.msoFalse,
                    WithWindow:=Microsoft.Office.Core.MsoTriState.msoFalse)

                Dim contentBuilder As New StringBuilder()
                contentBuilder.AppendLine($"文件: {Path.GetFileName(filePath)}")
                contentBuilder.AppendLine($"共 {presentation.Slides.Count} 张幻灯片")
                contentBuilder.AppendLine()

                ' 限制处理的幻灯片数量
                Dim maxSlides As Integer = Math.Min(presentation.Slides.Count, 20)

                For slideIndex As Integer = 1 To maxSlides
                    Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
                    Dim shapes As Microsoft.Office.Interop.PowerPoint.Shapes = Nothing
                    Try
                        slide = presentation.Slides(slideIndex)
                        contentBuilder.AppendLine($"=== 幻灯片 {slideIndex} ===")

                        ' 遍历幻灯片中的形状
                        shapes = slide.Shapes
                        For shapeIndex As Integer = 1 To shapes.Count
                            Dim shape As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
                            Try
                                shape = shapes(shapeIndex)

                                ' 检查是否有文本框架
                                If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                    If shape.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                        Dim textRange As Microsoft.Office.Interop.PowerPoint.TextRange = Nothing
                                        Try
                                            textRange = shape.TextFrame.TextRange
                                            Dim text As String = textRange.Text.Trim()
                                            If Not String.IsNullOrEmpty(text) Then
                                                ' 判断形状类型
                                                Dim shapeType As String = "文本"
                                                If shape.PlaceholderFormat IsNot Nothing Then
                                                    Select Case shape.PlaceholderFormat.Type
                                                        Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle,
                                                             Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle
                                                            shapeType = "标题"
                                                        Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle
                                                            shapeType = "副标题"
                                                        Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderBody
                                                            shapeType = "正文"
                                                    End Select
                                                End If
                                                contentBuilder.AppendLine($"  [{shapeType}] {text}")
                                            End If
                                        Finally
                                            ComObjectHelper.ReleaseComObject(textRange)
                                        End Try
                                    End If
                                End If

                                ' 检查是否是表格
                                If shape.HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                    Dim table As Microsoft.Office.Interop.PowerPoint.Table = Nothing
                                    Try
                                        table = shape.Table
                                        contentBuilder.AppendLine($"  [表格 {table.Rows.Count}行×{table.Columns.Count}列]")
                                        ' 读取表格内容（限制行数）
                                        Dim maxRows = Math.Min(table.Rows.Count, 10)
                                        For rowIdx = 1 To maxRows
                                            Dim rowContent As New StringBuilder("    ")
                                            For colIdx = 1 To table.Columns.Count
                                                Dim cell As Microsoft.Office.Interop.PowerPoint.Cell = Nothing
                                                Dim cellShape As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
                                                Dim cellTextRange As Microsoft.Office.Interop.PowerPoint.TextRange = Nothing
                                                Try
                                                    cell = table.Cell(rowIdx, colIdx)
                                                    cellShape = cell.Shape
                                                    cellTextRange = cellShape.TextFrame.TextRange
                                                    Dim cellText = cellTextRange.Text.Trim()
                                                    If cellText.Length > 20 Then cellText = cellText.Substring(0, 17) & "..."
                                                    rowContent.Append(cellText & " | ")
                                                Catch
                                                Finally
                                                    ComObjectHelper.ReleaseComObject(cellTextRange)
                                                    ComObjectHelper.ReleaseComObject(cellShape)
                                                    ComObjectHelper.ReleaseComObject(cell)
                                                End Try
                                            Next
                                            contentBuilder.AppendLine(rowContent.ToString().TrimEnd(" |".ToCharArray()))
                                        Next
                                    Finally
                                        ComObjectHelper.ReleaseComObject(table)
                                    End Try
                                End If
                            Catch shapeEx As Exception
                                Debug.WriteLine($"处理形状时出错: {shapeEx.Message}")
                            Finally
                                ComObjectHelper.ReleaseComObject(shape)
                            End Try
                        Next
                    Finally
                        ComObjectHelper.ReleaseComObject(shapes)
                        ComObjectHelper.ReleaseComObject(slide)
                    End Try

                    contentBuilder.AppendLine()
                Next

                If presentation.Slides.Count > maxSlides Then
                    contentBuilder.AppendLine($"... 共 {presentation.Slides.Count} 张幻灯片，仅显示前 {maxSlides} 张")
                End If

                Return New FileContentResult With {
                    .FileName = Path.GetFileName(filePath),
                    .FileType = "PowerPoint",
                    .ParsedContent = contentBuilder.ToString(),
                    .RawData = Nothing
                }

            Finally
                If presentation IsNot Nothing Then
                    presentation.Close()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(presentation)
                    GC.Collect()
                    GC.WaitForPendingFinalizers()
                End If
                ' 不调用 pptApp.Quit()，复用的是用户当前 PPT 进程
            End Try
        Catch ex As Exception
            Debug.WriteLine($"解析PowerPoint文件时出错: {ex.Message}")
            Return New FileContentResult With {
                .FileName = Path.GetFileName(filePath),
                .FileType = "PowerPoint",
                .ParsedContent = $"[解析PowerPoint文件时出错: {ex.Message}]"
            }
        End Try
    End Function
    Protected Overrides Function AppendCurrentSelectedContent(message As String) As String
        Try
            ' 检查是否启用了选择功能
            If Not selectedCellChecked Then
                Return message
            End If

            ' 获取当前 PowerPoint 中的选择
            Dim selection = Globals.ThisAddIn.Application.ActiveWindow.Selection
            If selection Is Nothing Then
                Return message
            End If

            ' 创建内容构建器，格式化选中内容
            Dim contentBuilder As New StringBuilder()
            contentBuilder.AppendLine(vbCrLf & "--- 用户选中的 PowerPoint 内容 ---")

            ' 添加演示文稿信息
            Dim activePresentation = Globals.ThisAddIn.Application.ActivePresentation
            If activePresentation IsNot Nothing Then
                contentBuilder.AppendLine($"演示文稿: {Path.GetFileName(activePresentation.FullName)}")
                contentBuilder.AppendLine($"当前幻灯片: {Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex}")
            End If

            ' 根据选择类型处理内容
            Select Case selection.Type
                Case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes
                    ' 处理形状选择（包括表格）
                    Dim shapeRange = selection.ShapeRange
                    contentBuilder.AppendLine($"选择类型: 形状 (共 {shapeRange.Count} 个)")

                    For i = 1 To shapeRange.Count
                        contentBuilder.AppendLine($"形状 {i}:")

                        ' 检查是否是表格
                        If shapeRange(i).HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            Dim table = shapeRange(i).Table
                            contentBuilder.AppendLine($"  表格: {table.Rows.Count} 行 × {table.Columns.Count} 列")

                            ' 添加表格内容
                            Dim maxRows As Integer = Math.Min(table.Rows.Count, 20)
                            Dim maxCols As Integer = Math.Min(table.Columns.Count, 10)

                            ' 处理表格头部
                            Dim headerBuilder As New StringBuilder("  ")
                            Dim separatorBuilder As New StringBuilder("  ")

                            For col = 1 To maxCols
                                Try
                                    Dim cellText = table.Cell(1, col).Shape.TextFrame.TextRange.Text.Trim()
                                    ' 限制单元格文本长度
                                    If cellText.Length > 20 Then
                                        cellText = cellText.Substring(0, 17) & "..."
                                    End If

                                    If col > 1 Then
                                        headerBuilder.Append(" | ")
                                        separatorBuilder.Append("-+-")
                                    End If
                                    headerBuilder.Append(cellText)
                                    separatorBuilder.Append(New String("-"c, Math.Max(cellText.Length, 3)))
                                Catch ex As Exception
                                    If col > 1 Then
                                        headerBuilder.Append(" | ")
                                        separatorBuilder.Append("-+-")
                                    End If
                                    headerBuilder.Append("N/A")
                                    separatorBuilder.Append("---")
                                End Try
                            Next

                            contentBuilder.AppendLine(headerBuilder.ToString())
                            contentBuilder.AppendLine(separatorBuilder.ToString())

                            ' 处理表格数据行
                            For row = 2 To maxRows
                                Dim rowBuilder As New StringBuilder("  ")

                                For col = 1 To maxCols
                                    Try
                                        Dim cellText = table.Cell(row, col).Shape.TextFrame.TextRange.Text.Trim()
                                        ' 限制单元格文本长度
                                        If cellText.Length > 20 Then
                                            cellText = cellText.Substring(0, 17) & "..."
                                        End If

                                        If col > 1 Then
                                            rowBuilder.Append(" | ")
                                        End If
                                        rowBuilder.Append(cellText)
                                    Catch ex As Exception
                                        If col > 1 Then
                                            rowBuilder.Append(" | ")
                                        End If
                                        rowBuilder.Append("N/A")
                                    End Try
                                Next

                                contentBuilder.AppendLine(rowBuilder.ToString())
                            Next

                            ' 添加表格说明
                            If table.Rows.Count > maxRows Then
                                contentBuilder.AppendLine($"  ... 共有 {table.Rows.Count} 行，仅显示前 {maxRows} 行")
                            End If

                            If table.Columns.Count > maxCols Then
                                contentBuilder.AppendLine($"  ... 共有 {table.Columns.Count} 列，仅显示前 {maxCols} 列")
                            End If
                        ElseIf shapeRange(i).HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            ' 处理文本框
                            Dim textFrame = shapeRange(i).TextFrame
                            If textFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                Dim text = textFrame.TextRange.Text.Trim()
                                ' 限制文本长度
                                If text.Length > 500 Then
                                    contentBuilder.AppendLine($"  文本: {text.Substring(0, 500)}...")
                                    contentBuilder.AppendLine($"  [文本太长，仅显示前500个字符，总计: {text.Length}个字符]")
                                Else
                                    contentBuilder.AppendLine($"  文本: {text}")
                                End If
                            Else
                                contentBuilder.AppendLine("  [空文本框]")
                            End If
                        ElseIf shapeRange(i).Type = Microsoft.Office.Core.MsoShapeType.msoPicture Then
                            ' 处理图片
                            contentBuilder.AppendLine("  [图片]")
                            If shapeRange(i).AlternativeText <> "" Then
                                contentBuilder.AppendLine($"  替代文本: {shapeRange(i).AlternativeText}")
                            End If
                        Else
                            ' 其他类型的形状
                            contentBuilder.AppendLine($"  [形状类型: {shapeRange(i).Type}]")
                        End If

                        ' 在形状之间添加分隔线
                        contentBuilder.AppendLine("  ---")
                    Next

                Case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText
                    ' 处理文本选择
                    contentBuilder.AppendLine("选择类型: 文本")

                    Dim textRange = selection.TextRange
                    If textRange IsNot Nothing Then
                        Dim text = textRange.Text.Trim()
                        ' 限制文本长度
                        If text.Length > 1000 Then
                            contentBuilder.AppendLine(text.Substring(0, 1000) & "...")
                            contentBuilder.AppendLine($"[文本太长，仅显示前1000个字符，总计: {text.Length}个字符]")
                        Else
                            contentBuilder.AppendLine(text)
                        End If
                    Else
                        contentBuilder.AppendLine("[无法获取文本内容]")
                    End If

                Case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionSlides
                    ' 处理幻灯片选择
                    Dim slideRange = selection.SlideRange
                    contentBuilder.AppendLine($"选择类型: 幻灯片 (共 {slideRange.Count} 张)")

                    ' 限制处理的幻灯片数量
                    Dim maxSlides = Math.Min(slideRange.Count, 5)

                    For i = 1 To maxSlides
                        Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
                        Try
                            slide = slideRange(i)
                            contentBuilder.AppendLine($"幻灯片 {slide.SlideIndex}:")

                            Dim title = GetSlideTitle(slide)
                            If title <> "[无标题]" Then
                                contentBuilder.AppendLine($"  标题: {title}")
                            Else
                                contentBuilder.AppendLine("  [无标题]")
                            End If

                            contentBuilder.Append(GetSlideContent(slide))
                            contentBuilder.AppendLine("  ---")
                        Finally
                            ComObjectHelper.ReleaseComObject(slide)
                        End Try
                    Next

                    ' 如果有更多幻灯片未显示，添加提示
                    If slideRange.Count > maxSlides Then
                        contentBuilder.AppendLine($"[共选中 {slideRange.Count} 张幻灯片，仅显示前 {maxSlides} 张]")
                    End If

                Case Else
                    contentBuilder.AppendLine($"选择类型: 未知 ({selection.Type})")
                    contentBuilder.AppendLine("[无法识别的选择类型]")
            End Select

            contentBuilder.AppendLine("--- 选中内容结束 ---" & vbCrLf)

            ' 返回原始消息加上选中内容
            Return message & contentBuilder.ToString()

        Catch ex As Exception
            Debug.WriteLine($"处理PowerPoint选中内容时出错: {ex.Message}")
            Return message ' 出错时返回原始消息
        End Try
    End Function

    ' 处理形状选择（包括表格）
    Private Function GetSlideTitle(slide As Microsoft.Office.Interop.PowerPoint.Slide) As String
        Try
            ' 检查幻灯片是否有标题占位符
            Dim shapes As Microsoft.Office.Interop.PowerPoint.Shapes = Nothing
            Try
                shapes = slide.Shapes
                For i = 1 To shapes.Count
                    Dim shape As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
                    Dim textRange As Microsoft.Office.Interop.PowerPoint.TextRange = Nothing
                    Try
                        shape = shapes(i)
                        If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder Then
                            If shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle Then
                                If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                    textRange = shape.TextFrame.TextRange
                                    Return textRange.Text.Trim()
                                End If
                            End If
                        End If
                    Finally
                        ComObjectHelper.ReleaseComObject(textRange)
                        ComObjectHelper.ReleaseComObject(shape)
                    End Try
                Next
            Finally
                ComObjectHelper.ReleaseComObject(shapes)
            End Try

            ' 如果没有找到标题占位符，尝试查找任何可能的标题
            Try
                shapes = slide.Shapes
                For i = 1 To shapes.Count
                    Dim shape As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
                    Dim textRange As Microsoft.Office.Interop.PowerPoint.TextRange = Nothing
                    Try
                        shape = shapes(i)
                        If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            textRange = shape.TextFrame.TextRange
                            Dim text = textRange.Text.Trim()
                            If Not String.IsNullOrEmpty(text) AndAlso text.Length < 100 Then
                                Return text ' 假设第一个简短文本是标题
                            End If
                        End If
                    Finally
                        ComObjectHelper.ReleaseComObject(textRange)
                        ComObjectHelper.ReleaseComObject(shape)
                    End Try
                Next
            Finally
                ComObjectHelper.ReleaseComObject(shapes)
            End Try

            Return "[无标题]"
        Catch ex As Exception
            Debug.WriteLine($"获取幻灯片标题时出错: {ex.Message}")
            Return "[获取标题出错]"
        End Try
    End Function

    ' 获取幻灯片内容
    Private Function GetSlideContent(slide As Microsoft.Office.Interop.PowerPoint.Slide) As String
        Try
            Dim contentBuilder As New StringBuilder()
            Dim processedTextShapes As Integer = 0
            Dim maxTextShapes As Integer = 5 ' 限制每张幻灯片处理的文本形状数量

            ' 处理幻灯片上的形状
            Dim shapes As Microsoft.Office.Interop.PowerPoint.Shapes = Nothing
            Try
                shapes = slide.Shapes
                For i = 1 To shapes.Count
                    Dim shape As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
                    Dim textRange As Microsoft.Office.Interop.PowerPoint.TextRange = Nothing
                    Try
                        shape = shapes(i)
                        ' 跳过标题形状，因为已经单独处理过了
                        If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder AndAlso
                       shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle Then
                            Continue For
                        End If

                        ' 处理文本形状
                        If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue AndAlso
                       shape.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then

                            If processedTextShapes >= maxTextShapes Then
                                contentBuilder.AppendLine("  [更多文本内容未显示...]")
                                Exit For
                            End If

                            textRange = shape.TextFrame.TextRange
                            Dim text = textRange.Text.Trim()
                            If Not String.IsNullOrEmpty(text) Then
                                ' 限制文本长度
                                If text.Length > 200 Then
                                    contentBuilder.AppendLine($"  文本: {text.Substring(0, 200)}...")
                                Else
                                    contentBuilder.AppendLine($"  文本: {text}")
                                End If
                                processedTextShapes += 1
                            End If
                            ' 处理表格形状
                        ElseIf shape.HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            contentBuilder.AppendLine("  [包含表格]")
                            ' 处理图片形状
                        ElseIf shape.Type = Microsoft.Office.Core.MsoShapeType.msoPicture Then
                            contentBuilder.AppendLine("  [包含图片]")
                            If shape.AlternativeText <> "" Then
                                contentBuilder.AppendLine($"  图片说明: {shape.AlternativeText}")
                            End If
                            ' 处理图表形状
                        ElseIf shape.Type = Microsoft.Office.Core.MsoShapeType.msoChart Then
                            contentBuilder.AppendLine("  [包含图表]")
                            ' 处理SmartArt形状
                        ElseIf shape.Type = Microsoft.Office.Core.MsoShapeType.msoSmartArt Then
                            contentBuilder.AppendLine("  [包含SmartArt图形]")
                        End If
                    Finally
                        ComObjectHelper.ReleaseComObject(textRange)
                        ComObjectHelper.ReleaseComObject(shape)
                    End Try
                Next
            Finally
                ComObjectHelper.ReleaseComObject(shapes)
            End Try

            ' 如果没有找到任何内容
            If contentBuilder.Length = 0 Then
                Return "  [幻灯片无可提取的文本内容]"
            End If

            Return contentBuilder.ToString()
        Catch ex As Exception
            Debug.WriteLine($"获取幻灯片内容时出错: {ex.Message}")
            Return $"  [获取内容出错: {ex.Message}]"
        End Try
    End Function

    Protected Overrides Function GetCurrentWorkingDirectory() As String
        Try
            ' 获取当前活动演示文稿的路径
            If Globals.ThisAddIn.Application.ActivePresentation IsNot Nothing Then
                Return Globals.ThisAddIn.Application.ActivePresentation.Path
            End If
        Catch ex As Exception
            Debug.WriteLine($"获取当前工作目录时出错: {ex.Message}")
        End Try

        ' 如果无法获取演示文稿路径，则返回应用程序目录
        Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
    End Function


    Protected Overrides Sub CheckAndCompleteProcessingHook(_finalUuid As String, allPlainMarkdownBuffer As StringBuilder)
        ' 在 AI 操作完成时创建撤销点 - 使用 ErrorHandler 包装
        ErrorHandlerExtension.SafeExecute(
            Sub()
                UndoManagerExtension.CreateAIOperationUndoPoint(
                    "PowerPoint",
                    Globals.ThisAddIn.Application,
                    "AI操作",
                    "AI 生成内容")
            End Sub,
            "PowerPointAi.ChatControl",
            "创建撤销点")

        ' 尝试检测并生成幻灯片 - 使用 ErrorHandler 包装
        ErrorHandlerExtension.SafeExecute(
            Sub()
                If allPlainMarkdownBuffer IsNot Nothing AndAlso allPlainMarkdownBuffer.Length > 0 Then
                    Dim aiResponse As String = allPlainMarkdownBuffer.ToString()

                    ' 检测并生成幻灯片
                    If PptGenerationHandlerExtension.DetectSlideOutline(aiResponse) Then
                        System.Diagnostics.Debug.WriteLine("[ChatControl] 检测到幻灯片大纲")

                        ' 在生成幻灯片前创建撤销点
                        UndoManagerExtension.CreateAIOperationUndoPoint(
                            "PowerPoint",
                            Globals.ThisAddIn.Application,
                            "AI幻灯片生成",
                            "生成AI幻灯片")

                        PptGenerationHandlerExtension.TryGenerateSlides(aiResponse, Globals.ThisAddIn.Application)
                    End If
                End If
            End Sub,
            "PowerPointAi.ChatControl",
            "幻灯片生成",
            "生成幻灯片时出错，请检查大纲格式。")

        ' 调用基类处理续写模式
        MyBase.CheckAndCompleteProcessingHook(_finalUuid, allPlainMarkdownBuffer)
    End Sub

    ' ========== 续写功能 ==========

    Private _continuationService As PowerPointContinuationService
    Private _cachedContinuationContext As ContinuationContext ' 缓存续写上下文，用于多轮续写

    ''' <summary>
    ''' 触发续写 - 获取光标上下文并发送AI请求
    ''' </summary>
    Protected Overrides Sub HandleTriggerContinuation(jsonDoc As JObject)
        Try
            ' 提取参数
            Dim style As String = ""
            Dim isContinuationMode As Boolean = False

            If jsonDoc IsNot Nothing Then
                If jsonDoc("style") IsNot Nothing Then
                    style = jsonDoc("style").ToString()
                End If
                If jsonDoc("isContinuationMode") IsNot Nothing Then
                    isContinuationMode = jsonDoc("isContinuationMode").ToObject(Of Boolean)()
                End If
            End If

            ' 初始化续写服务
            If _continuationService Is Nothing Then
                _continuationService = New PowerPointContinuationService(Globals.ThisAddIn.Application)
            End If

            ' 检查是否可以续写
            If Not _continuationService.CanContinue() Then
                GlobalStatusStrip.ShowWarning("无法获取演示文稿信息，请确保文档已打开")
                Return
            End If

            Dim context As ContinuationContext

            ' 如果是续写模式的后续请求，并且有缓存的上下文，则复用
            If isContinuationMode AndAlso _cachedContinuationContext IsNot Nothing Then
                ' 多轮续写：使用缓存的上下文，但style作为新的调整要求
                context = _cachedContinuationContext
                GlobalStatusStrip.ShowInfo("继续续写...")
            Else
                ' 首次续写或非续写模式：重新获取上下文
                context = _continuationService.GetCursorContext(3, 3)
                If context Is Nothing Then
                    GlobalStatusStrip.ShowWarning("无法获取幻灯片上下文")
                    Return
                End If
                ' 缓存上下文
                _cachedContinuationContext = context
                GlobalStatusStrip.ShowInfo("正在分析上下文并生成续写内容...")
            End If

            ' 发送续写请求（带上风格参数）
            SendContinuationRequest(context, style)

        Catch ex As Exception
            Debug.WriteLine($"HandleTriggerContinuation 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"触发续写时出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用续写结果到PowerPoint幻灯片
    ''' </summary>
    Protected Overrides Sub HandleApplyContinuation(jsonDoc As JObject)
        Try
            Dim content As String = If(jsonDoc("content") IsNot Nothing, jsonDoc("content").ToString(), String.Empty)
            Dim positionStr As String = If(jsonDoc("position") IsNot Nothing, jsonDoc("position").ToString(), "current")

            If String.IsNullOrWhiteSpace(content) Then
                GlobalStatusStrip.ShowWarning("续写内容为空")
                Return
            End If

            ' 确保续写服务已初始化
            If _continuationService Is Nothing Then
                _continuationService = New PowerPointContinuationService(Globals.ThisAddIn.Application)
            End If

            ' 根据position参数确定插入位置
            Dim insertPos As ShareRibbon.InsertPosition
            Select Case positionStr.ToLower()
                Case "start"
                    insertPos = ShareRibbon.InsertPosition.DocumentStart ' 首页
                Case "end"
                    insertPos = ShareRibbon.InsertPosition.DocumentEnd ' 末页
                Case Else ' "current" 或默认
                    insertPos = ShareRibbon.InsertPosition.AtCursor ' 当前页
            End Select

            ' 插入续写内容
            _continuationService.InsertContinuation(content, insertPos)

            GlobalStatusStrip.ShowInfo("续写内容已插入幻灯片")

            ' 通知前端移除操作按钮
            Dim uuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            If Not String.IsNullOrEmpty(uuid) Then
                ExecuteJavaScriptAsyncJS($"removeContinuationActions({UtilsService.ToJsStringLiteral(uuid)});")
            End If

        Catch ex As Exception
            Debug.WriteLine($"HandleApplyContinuation 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"插入续写内容时出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用模板渲染结果到PowerPoint幻灯片
    ''' </summary>
    Protected Overrides Sub HandleApplyTemplateContent(jsonDoc As JObject)
        Try
            Dim content As String = If(jsonDoc("content") IsNot Nothing, jsonDoc("content").ToString(), String.Empty)
            Dim positionStr As String = If(jsonDoc("position") IsNot Nothing, jsonDoc("position").ToString(), "current")

            If String.IsNullOrWhiteSpace(content) Then
                GlobalStatusStrip.ShowWarning("模板内容为空")
                Return
            End If

            ' 确保续写服务已初始化（复用其插入逻辑）
            If _continuationService Is Nothing Then
                _continuationService = New PowerPointContinuationService(Globals.ThisAddIn.Application)
            End If

            ' 根据position参数确定插入位置
            Dim insertPos As ShareRibbon.InsertPosition
            Select Case positionStr.ToLower()
                Case "start"
                    insertPos = ShareRibbon.InsertPosition.DocumentStart ' 首页
                Case "end"
                    insertPos = ShareRibbon.InsertPosition.DocumentEnd ' 末页
                Case Else ' "current" 或默认
                    insertPos = ShareRibbon.InsertPosition.AtCursor ' 当前页
            End Select

            ' 插入模板内容
            _continuationService.InsertContinuation(content, insertPos)

            GlobalStatusStrip.ShowInfo("模板内容已插入幻灯片")

            ' 通知前端移除操作按钮
            Dim uuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            If Not String.IsNullOrEmpty(uuid) Then
                ExecuteJavaScriptAsyncJS($"removeTemplateActions({UtilsService.ToJsStringLiteral(uuid)});")
            End If

        Catch ex As Exception
            Debug.WriteLine($"HandleApplyTemplateContent 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"插入模板内容时出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 获取当前PowerPoint上下文快照（用于自动补全）
    ''' </summary>
    Protected Overrides Function GetContextSnapshot() As JObject
        Dim snapshot As New JObject()
        snapshot("appType") = "PowerPoint"

        Try
            Dim pres = Globals.ThisAddIn.Application.ActivePresentation
            If pres IsNot Nothing Then
                snapshot("slidesCount") = pres.Slides.Count

                ' 获取当前幻灯片信息
                Try
                    Dim currentSlide As Object = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
                    If currentSlide IsNot Nothing Then
                        snapshot("currentSlide") = CInt(CallByName(currentSlide, "SlideIndex", CallType.Get))
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"GetContextSnapshot 获取当前幻灯片失败: {ex.GetType().Name}: {ex.Message}")
                End Try
            End If

            ' 获取选中内容
            Dim selText = ""
            Try
                Dim sel = Globals.ThisAddIn.Application.ActiveWindow.Selection
                If sel.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText Then
                    selText = sel.TextRange.Text
                ElseIf sel.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes Then
                    For i = 1 To Math.Min(sel.ShapeRange.Count, 3)
                        Dim shape = sel.ShapeRange(i)
                        If shape.HasTextFrame AndAlso shape.TextFrame.HasText Then
                            selText &= shape.TextFrame.TextRange.Text & " "
                        End If
                    Next
                End If
            Catch
            End Try

            If selText.Length > 300 Then
                selText = selText.Substring(0, 300) & "..."
            End If
            snapshot("selection") = selText.Trim()

        Catch ex As Exception
            Debug.WriteLine($"GetContextSnapshot 出错: {ex.Message}")
        End Try

        Return snapshot
    End Function

    ''' <summary>
    ''' 捕获 Agent 启动用 PowerPoint 上下文快照
    ''' </summary>
    Protected Overrides Function CaptureOfficeContext(appType As String) As Agent.Context.OfficeContext
        Try
            Return New Context.PowerPointContextProvider(Globals.ThisAddIn.Application).GetContext()
        Catch ex As Exception
            Debug.WriteLine($"CaptureOfficeContext 出错: {ex.Message}")
            Return New Agent.Context.OfficeContext With {.AppType = appType}
        End Try
    End Function

    ''' <summary>
    ''' 重写保存设置方法，同步更新PPT补全管理器状态
    ''' </summary>
    Protected Overrides Sub HandleSaveSettings(jsonDoc As JObject)
        MyBase.HandleSaveSettings(jsonDoc)
        
        ' 同步更新PPT补全管理器的启用状态
        Try
            Dim enableAutocomplete As Boolean = If(jsonDoc("enableAutocomplete")?.Value(Of Boolean)(), False)
            PowerPointCompletionManager.Instance.Enabled = enableAutocomplete
            Debug.WriteLine($"[PPTChatControl] 补全设置已同步: Enabled={enableAutocomplete}")
        Catch ex As Exception
            Debug.WriteLine($"[PPTChatControl] 同步补全设置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 执行JSON命令（重写基类方法）- 带严格验证
    ''' </summary>
    Private Function ExecuteJsonCommandCore(jsonCode As String, preview As Boolean) As Boolean
        Try
            ' 预览模式下跳过自动执行（排版/校对模式的JSON用于预览，由用户手动点击应用）
            If IsInPreviewMode() Then
                Debug.WriteLine($"[PPTChatControl] 预览模式({GetCurrentResponseMode()})下跳过JSON命令自动执行")
                Return True ' 返回True表示"成功处理"，避免显示错误
            End If

            ' 修复 AI 可能产生的双重转义（\\\" → \"，即字面反斜杠+引号 → 只保留引号）
            ' 原因：AI 有时对 code 字段内的引号进行两次转义，导致执行器解析失败
            If jsonCode.Contains("\""") Then
                jsonCode = jsonCode.Replace("\""", """")
                Debug.WriteLine("[PPTChatControl] 已修复双重转义引号")
            End If

            ' 使用严格的结构验证
            Dim errorMessage As String = ""
            Dim normalizedJson As JToken = Nothing
            
            If Not PowerPointJsonCommandSchema.ValidateJsonStructure(jsonCode, errorMessage, normalizedJson) Then
                ' 格式验证失败
                Debug.WriteLine($"PPT JSON格式验证失败: {errorMessage}")
                Debug.WriteLine($"原始JSON: {jsonCode.Substring(0, Math.Min(200, jsonCode.Length))}...")
                
                ShareRibbon.GlobalStatusStrip.ShowWarning($"JSON格式不符合规范: {errorMessage}")
                Return False
            End If
            
            ' 验证通过，根据类型执行
            If normalizedJson.Type = JTokenType.Object Then
                Dim jsonObj = CType(normalizedJson, JObject)
                
                ' 命令数组格式
                If jsonObj("commands") IsNot Nothing Then
                    Return ExecutePPTCommandsArray(jsonObj("commands"), jsonCode, preview)
                End If
                
                ' 单命令格式
                Return ExecutePPTSingleCommand(jsonObj, jsonCode, preview)
            End If
            
            ShareRibbon.GlobalStatusStrip.ShowWarning("无效的JSON格式")
            Return False

        Catch ex As Newtonsoft.Json.JsonReaderException
            ShareRibbon.GlobalStatusStrip.ShowWarning($"JSON格式无效: {ex.Message}")
            Return False
        Catch ex As Exception
            ShareRibbon.GlobalStatusStrip.ShowWarning($"执行失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 为 Agent 原生 PPT 工具返回真实 ToolResult，并以演示文稿/目标页快照生成最小 Observation。
    ''' </summary>
    Protected Overrides Function ExecuteJsonCommandWithToolResult(jsonCode As String, preview As Boolean) As Agent.ToolResult
        Dim envelope As JObject = Nothing
        Dim toolId As String = "PowerPointCommands"
        Try
            envelope = ParsePowerPointCommandEnvelope(jsonCode)
            Dim commandArray = TryCast(envelope?("commands"), JArray)
            If commandArray IsNot Nothing AndAlso commandArray.Count > 1 AndAlso
               commandArray.OfType(Of JObject)().Any(Function(item)
                   Return String.Equals(item("command")?.ToString(), "CreateSlides", StringComparison.OrdinalIgnoreCase)
               End Function) Then
                Return Agent.ToolResult.Failed("CreateSlides",
                                               "Professional CreateSlides must be the only command in its envelope",
                                               errorCode:=ExceptionClassifier.CodeOperationSchemaInvalid,
                                               userMessage:="专业整套幻灯片生成必须作为独立 CreateSlides 调用，请将后续图片、图表或对象操作拆成下一步",
                                               recoverable:=True,
                                               observation:=New JObject From {
                                                   {"kind", "write"},
                                                   {"summary", "CreateSlides 未执行：检测到混合命令数组"},
                                                   {"changed", False},
                                                   {"warnings", New JArray("mixed_create_slides_envelope")}
                                               })
            End If
            toolId = GetPowerPointEnvelopeToolId(envelope)
            If String.Equals(toolId, "DiscoverOfficeCapability", StringComparison.OrdinalIgnoreCase) Then
                Return OfficeRuntime.PowerPointApiCatalogProvider.SearchAsToolResult(
                    GetPowerPointEnvelopeParams(envelope))
            End If
            If String.Equals(toolId, "CreateSlides", StringComparison.OrdinalIgnoreCase) Then
                Return Design.ProfessionalDeckExecutor.ExecuteAsToolResult(
                    GetPowerPointEnvelopeParams(envelope), preview)
            End If
            If String.Equals(toolId, "OfficeObjectOperation", StringComparison.OrdinalIgnoreCase) Then
                Return OfficeRuntime.PowerPointOperationExecutor.Execute(
                    GetPowerPointEnvelopeParams(envelope))
            End If
            Dim beforeSnapshot = CapturePowerPointCommandSnapshot(envelope)
            Dim success = ExecuteJsonCommandCore(jsonCode, preview)
            ' 预览被用户拒绝时核心返回 True（UI 语义：取消不是错误），但 ToolResult 不能据此报成功。
            If ConsumePreviewCancelled() Then
                Return Agent.ToolResult.PreviewCancelled(toolId)
            End If
            Dim afterSnapshot = CapturePowerPointCommandSnapshot(envelope)
            Dim observation = BuildPowerPointCommandObservation(toolId, success, beforeSnapshot, afterSnapshot)
            Dim summary = If(observation?("summary")?.ToString(), $"{toolId} 执行完成")
            Dim observationError = If(afterSnapshot?("captureError")?.ToString(),
                                      beforeSnapshot?("captureError")?.ToString())

            ' 快照采集失败时 changed 恒为 False，成功回执其实没有任何证据支撑；
            ' 更不能让 Loop 把它当成“无变化”去重放，重放会在已改过的文稿上再改一次。
            If success AndAlso Not String.IsNullOrWhiteSpace(observationError) Then
                observation("rollback") = BuildPowerPointRollbackNote(True)
                Return Agent.ToolResult.Failed(
                    toolId,
                    $"PowerPoint 命令可能已执行，但结果观察失败: {observationError}",
                    errorCode:=ExceptionClassifier.CodeObservationFailed,
                    userMessage:="PowerPoint 操作可能已经完成，但插件无法安全验证结果；已停止自动重试",
                    recoverable:=False,
                    observation:=observation)
            End If

            If success Then
                Return Agent.ToolResult.Succeed(toolId,
                                                summary,
                                                data:=New With {.targetRefs = observation("targetRefs")},
                                                observation:=observation)
            End If

            ' 通用 PPT 命令没有逐命令事务边界（只有 CreateSlides 链按标记删页）：
            ' 失败时若文稿已经变化，只能如实上报半应用并掐断自动重放。
            Dim partialApply = observation.Value(Of Boolean)("changed")
            observation("rollback") = BuildPowerPointRollbackNote(partialApply)
            Dim failureMessage = If(partialApply, $"{summary}；文稿已被部分修改且未回滚", summary)
            Return Agent.ToolResult.Failed(toolId,
                                           failureMessage,
                                           errorCode:=If(partialApply,
                                                         ExceptionClassifier.CodePartialApply,
                                                         ExceptionClassifier.CodeUnknown),
                                           userMessage:=If(partialApply,
                                                           "PowerPoint 操作部分完成，插件无法自动撤销，请检查文稿后再继续",
                                                           failureMessage),
                                           recoverable:=Not partialApply,
                                           observation:=observation)
        Catch ex As Exception
            Return Agent.ToolResult.FromException(toolId, ex)
        End Try
    End Function

    ''' <summary>
    ''' 通用 PPT 命令不会自动撤销：attempted=False 表示根本没尝试回滚，
    ''' 避免 Agent 以为文稿已还原而继续规划（CreateSlides 链有自己的真实回滚）。
    ''' </summary>
    Private Shared Function BuildPowerPointRollbackNote(partialApply As Boolean) As JObject
        Return New JObject From {
            {"succeeded", False},
            {"attempted", False},
            {"verified", False},
            {"partialApply", partialApply},
            {"error", If(partialApply, "通用 PowerPoint 命令不支持安全地自动撤销本轮改动，请人工检查或手动撤销", "本轮未检测到文稿变化，无需回滚")}
        }
    End Function

    Private Shared Function ParsePowerPointCommandEnvelope(jsonCode As String) As JObject
        If String.IsNullOrWhiteSpace(jsonCode) Then Return Nothing
        Dim normalized = jsonCode
        Dim escapedQuote = ChrW(92) & ChrW(34)
        If normalized.Contains(escapedQuote) Then normalized = normalized.Replace(escapedQuote, ChrW(34))
        Return TryCast(JToken.Parse(normalized), JObject)
    End Function

    Private Shared Function GetPowerPointEnvelopeToolId(envelope As JObject) As String
        If envelope Is Nothing Then Return "PowerPointCommands"
        Dim commands = TryCast(envelope("commands"), JArray)
        If commands IsNot Nothing Then
            If commands.Count = 1 AndAlso commands(0).Type = JTokenType.Object Then
                Dim singleCommand = commands(0)("command")?.ToString()
                If Not String.IsNullOrWhiteSpace(singleCommand) Then Return singleCommand.Trim()
            End If
            Return "PowerPointCommands"
        End If
        Dim command = envelope("command")?.ToString()
        Return If(String.IsNullOrWhiteSpace(command), "PowerPointCommands", command.Trim())
    End Function

    Private Shared Function GetPowerPointEnvelopeParams(envelope As JObject) As JObject
        If envelope Is Nothing Then Return Nothing
        Dim commands = TryCast(envelope("commands"), JArray)
        If commands IsNot Nothing AndAlso commands.Count = 1 AndAlso commands(0).Type = JTokenType.Object Then
            Return TryCast(commands(0)("params"), JObject)
        End If
        Return TryCast(envelope("params"), JObject)
    End Function

    Private Function CapturePowerPointCommandSnapshot(envelope As JObject) As JObject
        Dim snapshot As New JObject()
        Try
            Dim app = Globals.ThisAddIn.Application
            Dim presentation = app.ActivePresentation
            If presentation Is Nothing Then Return snapshot

            snapshot("presentation") = If(presentation.Name, "")
            snapshot("slideCount") = CInt(presentation.Slides.Count)

            Dim slideIndex = ResolvePowerPointTargetSlideIndex(envelope, presentation.Slides.Count)
            If slideIndex <= 0 Then
                Try
                    slideIndex = CInt(app.ActiveWindow.View.Slide.SlideIndex)
                Catch
                    If presentation.Slides.Count > 0 Then slideIndex = 1
                End Try
            End If
            snapshot("slideIndex") = slideIndex

            If slideIndex > 0 AndAlso slideIndex <= presentation.Slides.Count Then
                Dim slide = presentation.Slides(slideIndex)
                snapshot("shapeCount") = CInt(slide.Shapes.Count)
                snapshot("titleHash") = ComputePowerPointObservationHash(GetPowerPointSlideTitle(slide))
                snapshot("textHash") = ComputePowerPointObservationHash(GetPowerPointSlideText(slide))
                snapshot("textPreview") = TruncatePowerPointObservationText(GetPowerPointSlideText(slide), 240)
                snapshot("notesHash") = ComputePowerPointObservationHash(GetPowerPointNotesText(slide))
                ' 动画与形状格式不在文本里：补齐这两类，否则 AddAnimation/FormatSlide 会被判为"无变化"而触发重放。
                snapshot("animationCount") = GetPowerPointAnimationCount(slide)
                snapshot("formatHash") = ComputePowerPointObservationHash(GetPowerPointFormatSignature(slide))
            End If
        Catch ex As Exception
            snapshot("captureError") = AppLogger.Redact(ex.Message)
        End Try
        Return snapshot
    End Function

    Private Shared Function ResolvePowerPointTargetSlideIndex(envelope As JObject, slideCount As Integer) As Integer
        If envelope Is Nothing Then Return 0
        Dim command = envelope
        If envelope("commands") IsNot Nothing AndAlso envelope("commands").Type = JTokenType.Array Then
            command = TryCast(envelope("commands").FirstOrDefault(), JObject)
        End If
        Dim rawIndex = command?("params")?("slideIndex")?.Value(Of Integer)()
        If rawIndex.HasValue Then
            ' 现有 PPT 命令参数大多按 0-based 解释，快照转换为 Office 1-based。
            Return Math.Max(1, Math.Min(slideCount, rawIndex.Value + 1))
        End If
        Return 0
    End Function

    Private Shared Function BuildPowerPointCommandObservation(toolId As String,
                                                              success As Boolean,
                                                              beforeSnapshot As JObject,
                                                              afterSnapshot As JObject) As JObject
        Dim changed = Not JToken.DeepEquals(beforeSnapshot, afterSnapshot)
        Dim slideIndex = If(afterSnapshot?("slideIndex")?.Value(Of Integer)(), beforeSnapshot?("slideIndex")?.Value(Of Integer)())
        Dim targetRefs As New JArray()
        If slideIndex > 0 Then targetRefs.Add($"PowerPoint:Slide/{slideIndex}") Else targetRefs.Add("PowerPoint:Presentation")

        Dim warnings As New JArray()
        If success AndAlso Not changed Then warnings.Add("命令已处理，但幻灯片快照未检测到变化；可能为用户取消、格式等价或 noop")
        If afterSnapshot?("captureError") IsNot Nothing Then warnings.Add(afterSnapshot("captureError"))

        Dim diff As New JObject From {
            {"slideCountDelta", GetPowerPointSnapshotInteger(afterSnapshot, "slideCount") - GetPowerPointSnapshotInteger(beforeSnapshot, "slideCount")},
            {"shapeCountDelta", GetPowerPointSnapshotInteger(afterSnapshot, "shapeCount") - GetPowerPointSnapshotInteger(beforeSnapshot, "shapeCount")},
            {"animationCountDelta", GetPowerPointSnapshotInteger(afterSnapshot, "animationCount") - GetPowerPointSnapshotInteger(beforeSnapshot, "animationCount")},
            {"formatChanged", Not String.Equals(GetPowerPointSnapshotString(beforeSnapshot, "formatHash"), GetPowerPointSnapshotString(afterSnapshot, "formatHash"), StringComparison.Ordinal)}
        }

        Return New JObject From {
            {"kind", "write"},
            {"summary", If(success, $"PowerPoint 工具 {toolId} 已执行", $"PowerPoint 工具 {toolId} 执行失败")},
            {"targetRefs", targetRefs},
            {"changed", changed},
            {"before", beforeSnapshot},
            {"after", afterSnapshot},
            {"diff", diff},
            {"warnings", warnings}
        }
    End Function

    Private Shared Function GetPowerPointSnapshotInteger(snapshot As JObject, name As String) As Integer
        If snapshot Is Nothing Then Return 0
        Return If(snapshot(name)?.Value(Of Integer)(), 0)
    End Function

    Private Shared Function GetPowerPointSnapshotString(snapshot As JObject, name As String) As String
        If snapshot Is Nothing Then Return ""
        Return If(snapshot(name)?.ToString(), "")
    End Function

    Private Shared Function ComputePowerPointObservationHash(value As String) As String
        Dim bytes = Encoding.UTF8.GetBytes(If(value, ""))
        Using sha = System.Security.Cryptography.SHA256.Create()
            Return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant()
        End Using
    End Function

    Private Shared Function GetPowerPointSlideTitle(slide As Object) As String
        Try
            If slide.Shapes.Title IsNot Nothing AndAlso slide.Shapes.Title.HasTextFrame Then
                Return If(slide.Shapes.Title.TextFrame.TextRange.Text, "")
            End If
        Catch
        End Try
        Return ""
    End Function

    Private Shared Function GetPowerPointSlideText(slide As Object) As String
        Dim parts As New List(Of String)()
        Try
            For Each shape As Object In slide.Shapes
                Try
                    If shape.HasTextFrame AndAlso shape.TextFrame.HasText Then
                        parts.Add(If(shape.TextFrame.TextRange.Text, ""))
                    End If
                Catch
                End Try
            Next
        Catch
        End Try
        Return String.Join(vbLf, parts)
    End Function

    Private Shared Function GetPowerPointNotesText(slide As Object) As String
        Dim parts As New List(Of String)()
        Try
            For Each shape As Object In slide.NotesPage.Shapes
                Try
                    If shape.HasTextFrame AndAlso shape.TextFrame.HasText Then parts.Add(If(shape.TextFrame.TextRange.Text, ""))
                Catch
                End Try
            Next
        Catch
        End Try
        Return String.Join(vbLf, parts)
    End Function

    Private Shared Function GetPowerPointAnimationCount(slide As Object) As Integer
        Try
            Dim mainSeq = slide.TimeLine.MainSequence
            If mainSeq IsNot Nothing Then Return CInt(mainSeq.Count)
        Catch
        End Try
        Return 0
    End Function

    Private Shared Function GetPowerPointFormatSignature(slide As Object) As String
        Dim parts As New List(Of String)()
        Try
            For Each shape As Object In slide.Shapes
                Try
                    Dim segment As New StringBuilder()
                    segment.Append(shape.Name).Append("|")
                    segment.Append(CInt(shape.Left)).Append(",")
                    segment.Append(CInt(shape.Top)).Append(",")
                    segment.Append(CInt(shape.Width)).Append(",")
                    segment.Append(CInt(shape.Height))
                    Try
                        If CType(shape.Fill.Type, Integer) = 1 Then segment.Append(";fill=").Append(shape.Fill.ForeColor.RGB)
                    Catch
                    End Try
                    If shape.HasTextFrame AndAlso shape.TextFrame.HasText Then
                        Dim tr = shape.TextFrame.TextRange.Font
                        segment.Append(";font=").Append(If(tr.Name, "")).Append(",")
                        segment.Append(CSng(tr.Size)).Append(",b=").Append(CInt(tr.Bold)).Append(",i=").Append(CInt(tr.Italic))
                        Try
                            segment.Append(",color=").Append(tr.Color.RGB)
                        Catch
                        End Try
                    End If
                    parts.Add(segment.ToString())
                Catch
                End Try
            Next
        Catch
        End Try
        Return String.Join(vbLf, parts)
    End Function

    Private Shared Function TruncatePowerPointObservationText(value As String, maxLength As Integer) As String
        Dim text = If(value, "").Replace(vbCr, " ").Replace(vbLf, " ").Trim()
        If text.Length <= maxLength Then Return text
        Return text.Substring(0, maxLength)
    End Function

    ''' <summary>
    ''' 执行PPT命令数组
    ''' </summary>
    Private Function ExecutePPTCommandsArray(commandsArray As JToken, originalJson As String, preview As Boolean) As Boolean
        Try
            Dim commands = CType(commandsArray, JArray)
            If commands.Count = 0 Then
                ShareRibbon.GlobalStatusStrip.ShowWarning("命令数组为空")
                Return False
            End If

            ' 预览所有命令 - 使用增强的预览表单
            If preview Then
                If Not ShareRibbon.CommandPreviewForm.ShowPreview($"PPT命令预览 - 共 {commands.Count} 个命令", commandsArray) Then
                    ExecuteJavaScriptAsyncJS("handleExecutionCancelled('')")
                    NotifyPreviewCancelled()
                    Return True
                End If
            End If

            ' 执行所有命令
            Dim successCount = 0
            Dim failCount = 0

            For Each cmd In commands
                If cmd.Type = JTokenType.Object Then
                    Dim cmdObj = CType(cmd, JObject)
                    If ExecutePPTCommand(cmdObj, preview) Then
                        successCount += 1
                    Else
                        failCount += 1
                    End If
                End If
            Next

            If failCount = 0 Then
                ShareRibbon.GlobalStatusStrip.ShowInfo($"所有 {successCount} 个命令执行成功")
            Else
                ShareRibbon.GlobalStatusStrip.ShowWarning($"执行完成: {successCount} 成功, {failCount} 失败")
            End If

            Return failCount = 0

        Catch ex As Exception
            Debug.WriteLine($"ExecutePPTCommandsArray 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"批量执行失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行单个PPT命令
    ''' </summary>
    Private Function ExecutePPTSingleCommand(commandJson As JObject, processedJson As String, preview As Boolean) As Boolean
        Try
            Dim command = commandJson("command")?.ToString()
            
            ' 预览 - 使用增强的预览表单
            If preview Then
                If Not ShareRibbon.CommandPreviewForm.ShowPreview("PPT命令预览", commandJson) Then
                    ExecuteJavaScriptAsyncJS("handleExecutionCancelled('')")
                    NotifyPreviewCancelled()
                    Return True
                End If
            End If

            ' 执行命令
            Dim success = ExecutePPTCommand(commandJson, preview)

            If success Then
                ShareRibbon.GlobalStatusStrip.ShowInfo($"命令 '{command}' 执行成功")
            Else
                ShareRibbon.GlobalStatusStrip.ShowWarning($"命令 '{command}' 执行失败")
            End If

            Return success

        Catch ex As Exception
            Debug.WriteLine($"ExecutePPTSingleCommand 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行具体的PPT命令
    ''' </summary>
    Private Function ExecutePPTCommand(commandJson As JObject, Optional preview As Boolean = False) As Boolean
        Try
            Dim command = commandJson("command")?.ToString()
            Dim params = commandJson("params")
            
            Dim pres = Globals.ThisAddIn.Application.ActivePresentation

            Select Case command.ToLower()
                Case "insertslide"
                    Return ExecuteInsertSlide(params, pres)
                Case "inserttext"
                    Return ExecuteInsertText(params, pres)
                Case "insertshape"
                    Return ExecuteInsertShape(params, pres)
                Case "formatslide"
                    Return ExecuteFormatSlide(params, pres)
                Case "inserttable"
                    Return ExecuteInsertTable(params, pres)
                Case "createslides"
                    If preview AndAlso params IsNot Nothing Then params("preview") = True
                    Return ExecuteCreateSlides(params, pres)
                Case "addanimation"
                    Return ExecuteAddAnimation(params, pres)
                Case "applytransition"
                    Return ExecuteApplyTransition(params, pres)
                Case "beautifyslides"
                    Return ExecuteBeautifySlides(params, pres)
                Case "deleteslide"
                    Return ExecuteDeleteSlide(params, pres)
                Case "duplicateslide"
                    Return ExecuteDuplicateSlide(params, pres)
                Case "moveslide"
                    Return ExecuteMoveSlide(params, pres)
                Case "setslidelayout"
                    Return ExecuteSetSlideLayout(params, pres)
                Case "applytheme"
                    Return ExecuteApplyTheme(params, pres)
                Case "addspeakernotes"
                    Return ExecuteAddSpeakerNotes(params, pres)
                Case "executevba"
                    Dim vbaCode = params("code")?.ToString()
                    If String.IsNullOrEmpty(vbaCode) Then
                        GlobalStatusStrip.ShowWarning("ExecuteVBA 缺少 code 参数")
                        Return False
                    End If
                    Return CodeExecutionService.ExecuteVBACode(vbaCode, False)
                Case Else
                    Debug.WriteLine($"不支持的PPT命令: {command}")
                    GlobalStatusStrip.ShowWarning($"暂不支持的PPT命令: {command}")
                    Return False
            End Select

        Catch ex As Exception
            Debug.WriteLine($"ExecutePPTCommand 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteInsertSlide(params As JToken, pres As Object) As Boolean
        Try
            Dim position = If(params("position")?.ToString(), "end")
            Dim title = If(params("title")?.ToString(), "")
            Dim content = If(params("content")?.ToString(), "")

            Dim slideIndex As Integer
            If position.ToLower() = "end" Then
                slideIndex = pres.Slides.Count + 1
            ElseIf position.ToLower() = "current" Then
                slideIndex = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex + 1
            Else
                slideIndex = pres.Slides.Count + 1
            End If

            ' 添加幻灯片 (使用标题和内容布局 ppLayoutTitleOnly = 11)
            Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
            Try
                slide = pres.Slides.Add(slideIndex, 11)

                TrySetPlaceholderText(slide, title, Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle)

                ' 如果有内容，添加文本框
                If Not String.IsNullOrEmpty(content) Then
                    AddTextBoxToSlide(slide, content, 50, 150, 600, 300)
                End If
            Finally
                ComObjectHelper.ReleaseComObject(slide)
            End Try

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteInsertSlide 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteInsertText(params As JToken, pres As Object) As Boolean
        Try
            Dim content = params("content")?.ToString()
            Dim slideIndex = If(params("slideIndex")?.Value(Of Integer)(), -1)
            Dim x = If(params("x")?.Value(Of Single)(), 100)
            Dim y = If(params("y")?.Value(Of Single)(), 200)

            Dim slide As Object
            If slideIndex < 0 Then
                slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
            Else
                slide = pres.Slides(Math.Min(slideIndex + 1, pres.Slides.Count))
            End If

            Dim textBox = slide.Shapes.AddTextbox(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                x, y, 400, 100)
            textBox.TextFrame.TextRange.Text = content

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteInsertText 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteInsertShape(params As JToken, pres As Object) As Boolean
        Try
            Dim shapeType = If(params("shapeType")?.ToString(), "rectangle")
            Dim x = params("x")?.Value(Of Single)()
            Dim y = params("y")?.Value(Of Single)()
            Dim width = If(params("width")?.Value(Of Single)(), 100)
            Dim height = If(params("height")?.Value(Of Single)(), 100)

            Dim slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide

            ' 根据shapeType添加不同形状
            Dim msoShapeType As Integer = 1 ' msoShapeRectangle
            Select Case shapeType.ToLower()
                Case "rectangle"
                    msoShapeType = 1
                Case "oval", "circle"
                    msoShapeType = 9 ' msoShapeOval
                Case "triangle"
                    msoShapeType = 7 ' msoShapeIsoscelesTriangle
                Case "arrow"
                    msoShapeType = 13 ' msoShapeRightArrow
            End Select

            slide.Shapes.AddShape(msoShapeType, x, y, width, height)
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteInsertShape 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteFormatSlide(params As JToken, pres As Object) As Boolean
        Try
            Dim slideIndex = If(params("slideIndex")?.Value(Of Integer)(), -1)
            
            Dim slide As Object
            If slideIndex < 0 Then
                slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
            Else
                slide = pres.Slides(Math.Min(slideIndex + 1, pres.Slides.Count))
            End If

            ' 设置背景
            Dim background = params("background")?.ToString()
            If Not String.IsNullOrEmpty(background) Then
                Try
                    ' 尝试解析颜色
                    Dim color = System.Drawing.ColorTranslator.FromHtml(background)
                    slide.FollowMasterBackground = False
                    slide.Background.Fill.Solid()
                    slide.Background.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(color)
                Catch
                End Try
            End If

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteFormatSlide 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteInsertTable(params As JToken, pres As Object) As Boolean
        Try
            Dim rows = params("rows")?.Value(Of Integer)()
            Dim cols = params("cols")?.Value(Of Integer)()
            Dim slideIndex = If(params("slideIndex")?.Value(Of Integer)(), -1)

            If rows <= 0 OrElse cols <= 0 Then Return False

            Dim slide As Object
            If slideIndex < 0 Then
                slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
            Else
                slide = pres.Slides(Math.Min(slideIndex + 1, pres.Slides.Count))
            End If

            Dim table = slide.Shapes.AddTable(rows, cols, 50, 150, 600, 300)

            ' 如果有data，填充表格
            Dim data = params("data")
            If data IsNot Nothing AndAlso data.Type = JTokenType.Array Then
                Dim dataArr = CType(data, JArray)
                Dim x As Integer = dataArr.Count - 1
                Dim x2 As Integer = rows - 1
                For rowIdx = 0 To Math.Min(x, x2)
                    Dim rowData = dataArr(rowIdx)
                    If rowData.Type = JTokenType.Array Then
                        Dim rowArr = CType(rowData, JArray)
                        Dim y As Integer = rowArr.Count - 1
                        Dim y1 As Integer = cols - 1
                        For colIdx = 0 To Math.Min(y, y1)
                            table.Table.Cell(rowIdx + 1, colIdx + 1).Shape.TextFrame.TextRange.Text = rowArr(colIdx).ToString()
                        Next
                    End If
                Next
            End If

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteInsertTable 出错: {ex.Message}")
            Return False
        End Try
    End Function

#Region "高级PPT命令实现"

    ''' <summary>
    ''' 批量创建幻灯片
    ''' </summary>
    Private Function ExecuteCreateSlides(params As JToken, pres As Object) As Boolean
        Try
            Dim slides = params("slides")
            If slides Is Nothing OrElse slides.Type <> JTokenType.Array Then
                Return False
            End If

            Dim slidesArray = CType(slides, JArray)
            Dim startIndex = pres.Slides.Count + 1

            For i = 0 To slidesArray.Count - 1
                Dim slideData = slidesArray(i)
                Dim title = If(slideData("title")?.ToString(), "")
                Dim content = If(slideData("content")?.ToString(), "")
                Dim layout = If(slideData("layout")?.ToString(), "titleAndContent")

                ' 根据layout选择布局类型
                Dim layoutType As Integer = GetLayoutType(layout)
                Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
                Try
                    slide = pres.Slides.Add(startIndex + i, layoutType)

                    TrySetPlaceholderText(slide, title, Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle)

                    ' 填充内容
                    If Not String.IsNullOrEmpty(content) Then
                        Dim contentFilled = TrySetPlaceholderText(
                            slide,
                            content,
                            Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderBody,
                            Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle)

                        ' 如果没有内容占位符，添加文本框
                        If Not contentFilled Then
                            AddTextBoxToSlide(slide, content, 50, 150, 620, 350, 18)
                        End If
                    End If

                    ' CreateSlides 不能把 Office 默认白底文本当成可交付设计。
                    BeautifySingleSlide(slide, BuildDefaultPptTheme(i))
                Finally
                    ComObjectHelper.ReleaseComObject(slide)
                End Try
            Next

            ShareRibbon.GlobalStatusStrip.ShowInfo($"成功创建 {slidesArray.Count} 张幻灯片")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteCreateSlides 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function TrySetPlaceholderText(
        slide As Microsoft.Office.Interop.PowerPoint.Slide,
        text As String,
        ParamArray placeholderTypes() As Microsoft.Office.Interop.PowerPoint.PpPlaceholderType) As Boolean

        If slide Is Nothing OrElse String.IsNullOrEmpty(text) Then Return False

        Dim shapes As Microsoft.Office.Interop.PowerPoint.Shapes = Nothing
        Try
            shapes = slide.Shapes
            For i = 1 To shapes.Count
                Dim shape As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
                Dim textRange As Microsoft.Office.Interop.PowerPoint.TextRange = Nothing
                Try
                    shape = shapes(i)
                    If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder Then
                        Dim placeholderType = shape.PlaceholderFormat.Type
                        If placeholderTypes.Contains(placeholderType) Then
                            textRange = shape.TextFrame.TextRange
                            textRange.Text = text
                            Return True
                        End If
                    End If
                Finally
                    ComObjectHelper.ReleaseComObject(textRange)
                    ComObjectHelper.ReleaseComObject(shape)
                End Try
            Next
        Finally
            ComObjectHelper.ReleaseComObject(shapes)
        End Try

        Return False
    End Function

    Private Sub AddTextBoxToSlide(
        slide As Microsoft.Office.Interop.PowerPoint.Slide,
        text As String,
        x As Single,
        y As Single,
        width As Single,
        height As Single,
        Optional fontSize As Single = 0)

        Dim shapes As Microsoft.Office.Interop.PowerPoint.Shapes = Nothing
        Dim textBox As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
        Dim textRange As Microsoft.Office.Interop.PowerPoint.TextRange = Nothing
        Try
            shapes = slide.Shapes
            textBox = shapes.AddTextbox(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                x, y, width, height)
            textRange = textBox.TextFrame.TextRange
            textRange.Text = text
            If fontSize > 0 Then
                textRange.Font.Size = fontSize
            End If
        Finally
            ComObjectHelper.ReleaseComObject(textRange)
            ComObjectHelper.ReleaseComObject(textBox)
            ComObjectHelper.ReleaseComObject(shapes)
        End Try
    End Sub

    ''' <summary>
    ''' 获取布局类型
    ''' </summary>
    Private Function GetLayoutType(layout As String) As Integer
        Select Case layout.ToLower()
            Case "title", "titleonly"
                Return 11 ' ppLayoutTitleOnly
            Case "titleandcontent", "content"
                Return 2 ' ppLayoutText
            Case "twocontent", "twotext"
                Return 3 ' ppLayoutTwoColumnText
            Case "blank"
                Return 7 ' ppLayoutBlank
            Case "sectionheader"
                Return 1 ' ppLayoutTitle
            Case Else
                Return 2 ' ppLayoutText (默认)
        End Select
    End Function

    ''' <summary>
    ''' 添加动画效果
    ''' </summary>
    Private Function ExecuteAddAnimation(params As JToken, pres As Object) As Boolean
        Try
            Dim slideIndex = If(params("slideIndex")?.Value(Of Integer)(), -1)
            Dim effect = If(params("effect")?.ToString(), "fadeIn")
            Dim targetShapes = If(params("targetShapes")?.ToString(), "all")

            Dim msoEffect = GetMsoAnimEffect(effect)
            Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
            Dim timeline As Microsoft.Office.Interop.PowerPoint.TimeLine = Nothing
            Dim sequence As Microsoft.Office.Interop.PowerPoint.Sequence = Nothing
            Dim shapes As Microsoft.Office.Interop.PowerPoint.Shapes = Nothing
            Try
                If slideIndex < 0 Then
                    slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
                Else
                    slide = pres.Slides(Math.Min(slideIndex + 1, pres.Slides.Count))
                End If

                timeline = slide.TimeLine
                sequence = timeline.MainSequence
                shapes = slide.Shapes

                For i = 1 To shapes.Count
                    Dim shape As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
                    Try
                        shape = shapes(i)
                        Dim shouldAnimate = False

                        If targetShapes.ToLower() = "all" Then
                            shouldAnimate = True
                        ElseIf targetShapes.ToLower() = "title" Then
                            If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder AndAlso
                               shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle Then
                                shouldAnimate = True
                            End If
                        ElseIf targetShapes.ToLower() = "content" Then
                            If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder AndAlso
                               (shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderBody OrElse
                                shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle) Then
                                shouldAnimate = True
                            End If
                        End If

                        If shouldAnimate AndAlso shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            Dim animationEffect As Microsoft.Office.Interop.PowerPoint.Effect = Nothing
                            Try
                                animationEffect = sequence.AddEffect(shape, msoEffect)
                            Catch
                                ' 某些形状可能不支持动画
                            Finally
                                ComObjectHelper.ReleaseComObject(animationEffect)
                            End Try
                        End If
                    Finally
                        ComObjectHelper.ReleaseComObject(shape)
                    End Try
                Next
            Finally
                ComObjectHelper.ReleaseComObject(shapes)
                ComObjectHelper.ReleaseComObject(sequence)
                ComObjectHelper.ReleaseComObject(timeline)
                ComObjectHelper.ReleaseComObject(slide)
            End Try

            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteAddAnimation 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 获取动画效果类型
    ''' </summary>
    Private Function GetMsoAnimEffect(effect As String) As Integer
        Select Case effect.ToLower()
            Case "fadein", "fade"
                Return 10 ' msoAnimEffectFade
            Case "flyin", "fly"
                Return 2 ' msoAnimEffectFly
            Case "zoom"
                Return 53 ' msoAnimEffectGrowAndTurn
            Case "wipe"
                Return 22 ' msoAnimEffectWipe
            Case "appear"
                Return 1 ' msoAnimEffectAppear
            Case "float"
                Return 42 ' msoAnimEffectFloat
            Case Else
                Return 10 ' msoAnimEffectFade (默认)
        End Select
    End Function

    ''' <summary>
    ''' 应用幻灯片切换效果
    ''' </summary>
    Private Function ExecuteApplyTransition(params As JToken, pres As Object) As Boolean
        Try
            Dim transType = If(params("transitionType")?.ToString(), "fade")
            Dim scope = If(params("scope")?.ToString(), "all")
            Dim duration = If(params("duration")?.Value(Of Single)(), 1.0F)

            Dim transEffect = GetTransitionEffect(transType)

            Dim processedCount As Integer = 0
            If scope.ToLower() = "all" Then
                Dim slides As Microsoft.Office.Interop.PowerPoint.Slides = Nothing
                Try
                    slides = pres.Slides
                    For i = 1 To slides.Count
                        Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
                        Dim transition As Microsoft.Office.Interop.PowerPoint.SlideShowTransition = Nothing
                        Try
                            slide = slides(i)
                            transition = slide.SlideShowTransition
                            transition.EntryEffect = transEffect
                            transition.Duration = duration
                            transition.AdvanceOnClick = True
                            processedCount += 1
                        Finally
                            ComObjectHelper.ReleaseComObject(transition)
                            ComObjectHelper.ReleaseComObject(slide)
                        End Try
                    Next
                Finally
                    ComObjectHelper.ReleaseComObject(slides)
                End Try
            Else
                Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
                Dim transition As Microsoft.Office.Interop.PowerPoint.SlideShowTransition = Nothing
                Try
                    slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
                    transition = slide.SlideShowTransition
                    transition.EntryEffect = transEffect
                    transition.Duration = duration
                    transition.AdvanceOnClick = True
                    processedCount = 1
                Finally
                    ComObjectHelper.ReleaseComObject(transition)
                    ComObjectHelper.ReleaseComObject(slide)
                End Try
            End If

            ShareRibbon.GlobalStatusStrip.ShowInfo($"已为 {processedCount} 张幻灯片应用切换效果")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteApplyTransition 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 获取切换效果类型
    ''' </summary>
    Private Function GetTransitionEffect(transType As String) As Integer
        Select Case transType.ToLower()
            Case "fade"
                Return 257 ' ppTransitionFade
            Case "push"
                Return 3844 ' ppTransitionPush
            Case "wipe"
                Return 769 ' ppTransitionWipe
            Case "split"
                Return 2817 ' ppTransitionSplit
            Case "reveal"
                Return 3073 ' ppTransitionReveal
            Case "random"
                Return 513 ' ppTransitionRandom
            Case Else
                Return 257 ' ppTransitionFade (默认)
        End Select
    End Function

    ''' <summary>
    ''' 美化幻灯片
    ''' </summary>
    Private Function ExecuteBeautifySlides(params As JToken, pres As Object) As Boolean
        Try
            Dim scope = If(params("scope")?.ToString(), "all")
            Dim theme = params("theme")

            Dim processedCount As Integer = 0
            If scope.ToLower() = "all" Then
                Dim slides As Microsoft.Office.Interop.PowerPoint.Slides = Nothing
                Try
                    slides = pres.Slides
                    For i = 1 To slides.Count
                        Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
                        Try
                            slide = slides(i)
                            BeautifySingleSlide(slide, theme)
                            processedCount += 1
                        Finally
                            ComObjectHelper.ReleaseComObject(slide)
                        End Try
                    Next
                Finally
                    ComObjectHelper.ReleaseComObject(slides)
                End Try
            Else
                Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = Nothing
                Try
                    slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
                    BeautifySingleSlide(slide, theme)
                    processedCount = 1
                Finally
                    ComObjectHelper.ReleaseComObject(slide)
                End Try
            End If

            ShareRibbon.GlobalStatusStrip.ShowInfo($"已美化 {processedCount} 张幻灯片")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteBeautifySlides 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Sub BeautifySingleSlide(slide As Microsoft.Office.Interop.PowerPoint.Slide, theme As JToken)
        If slide Is Nothing Then Return
        If theme Is Nothing Then theme = BuildDefaultPptTheme(Math.Max(0, slide.SlideIndex - 1))

        ' 应用背景色
        If theme IsNot Nothing AndAlso theme("background") IsNot Nothing Then
            Try
                Dim bgColor = theme("background").ToString()
                Dim color = System.Drawing.ColorTranslator.FromHtml(bgColor)
                slide.FollowMasterBackground = False
                slide.Background.Fill.Solid()
                slide.Background.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(color)
            Catch
            End Try
        End If

        ' 应用字体样式
        Dim shapes As Microsoft.Office.Interop.PowerPoint.Shapes = Nothing
        Try
            shapes = slide.Shapes
            For i = 1 To shapes.Count
                Dim shape As Microsoft.Office.Interop.PowerPoint.Shape = Nothing
                Dim textRange As Microsoft.Office.Interop.PowerPoint.TextRange = Nothing
                Try
                    shape = shapes(i)
                    If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                        Dim isTitle = False
                        If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder Then
                            isTitle = (shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle)
                        End If

                        Dim fontTheme = If(isTitle, theme?("titleFont"), theme?("bodyFont"))
                        If fontTheme IsNot Nothing Then
                            textRange = shape.TextFrame.TextRange
                            ApplyFontStyle(textRange, fontTheme)
                        End If
                    End If
                Finally
                    ComObjectHelper.ReleaseComObject(textRange)
                    ComObjectHelper.ReleaseComObject(shape)
                End Try
            Next
        Finally
            ComObjectHelper.ReleaseComObject(shapes)
        End Try
    End Sub

    Private Function BuildDefaultPptTheme(themeVariant As Integer) As JObject
        Dim backgrounds = {"#F7F3EA", "#EEF4F7", "#F5F1F7"}
        Dim accents = {"#7A3E2E", "#235B70", "#62406F"}
        Dim index = Math.Abs(themeVariant) Mod backgrounds.Length
        Return New JObject From {
            {"background", backgrounds(index)},
            {"titleFont", New JObject From {
                {"name", "Microsoft YaHei"},
                {"size", 30},
                {"bold", True},
                {"color", accents(index)}
            }},
            {"bodyFont", New JObject From {
                {"name", "Microsoft YaHei"},
                {"size", 20},
                {"color", "#27323A"}
            }}
        }
    End Function

    ''' <summary>
    ''' 应用字体样式
    ''' </summary>
    Private Sub ApplyFontStyle(textRange As Object, fontTheme As JToken)
        Try
            If fontTheme("name") IsNot Nothing Then
                textRange.Font.Name = fontTheme("name").ToString()
            End If
            If fontTheme("size") IsNot Nothing Then
                textRange.Font.Size = fontTheme("size").Value(Of Single)()
            End If
            If fontTheme("color") IsNot Nothing Then
                Dim colorStr = fontTheme("color").ToString()
                Dim color = System.Drawing.ColorTranslator.FromHtml(colorStr)
                textRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(color)
            End If
            If fontTheme("bold") IsNot Nothing Then
                textRange.Font.Bold = If(fontTheme("bold").Value(Of Boolean)(), -1, 0)
            End If
        Catch ex As Exception
            Debug.WriteLine($"ApplyFontStyle 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 删除幻灯片
    ''' </summary>
    Private Function ExecuteDeleteSlide(params As JToken, pres As Object) As Boolean
        Try
            Dim slideIndex As Integer = If(params("slideIndex") IsNot Nothing, params("slideIndex").Value(Of Integer)(), -1)
            If slideIndex = -1 OrElse slideIndex = 0 Then
                slideIndex = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex
            End If
            If pres.Slides.Count <= 1 Then
                GlobalStatusStrip.ShowWarning("无法删除：演示文稿必须至少保留一张幻灯片")
                Return False
            End If
            If slideIndex < 1 OrElse slideIndex > pres.Slides.Count Then
                GlobalStatusStrip.ShowWarning($"幻灯片索引 {slideIndex} 超出范围（共 {pres.Slides.Count} 张）")
                Return False
            End If
            pres.Slides(slideIndex).Delete()
            GlobalStatusStrip.ShowInfo($"已删除第 {slideIndex} 张幻灯片")
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteDeleteSlide 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 复制幻灯片
    ''' </summary>
    Private Function ExecuteDuplicateSlide(params As JToken, pres As Object) As Boolean
        Try
            Dim slideIndex As Integer = If(params("slideIndex") IsNot Nothing, params("slideIndex").Value(Of Integer)(), -1)
            If slideIndex = -1 OrElse slideIndex = 0 Then
                slideIndex = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex
            End If
            If slideIndex < 1 OrElse slideIndex > pres.Slides.Count Then Return False
            pres.Slides(slideIndex).Copy()
            pres.Slides.Paste(slideIndex + 1)
            GlobalStatusStrip.ShowInfo($"已复制第 {slideIndex} 张幻灯片")
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteDuplicateSlide 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 移动幻灯片
    ''' </summary>
    Private Function ExecuteMoveSlide(params As JToken, pres As Object) As Boolean
        Try
            Dim fromIndex As Integer = If(params("fromIndex") IsNot Nothing, params("fromIndex").Value(Of Integer)(), -1)
            Dim toIndex As Integer = If(params("toIndex") IsNot Nothing, params("toIndex").Value(Of Integer)(), -1)
            If fromIndex < 1 OrElse fromIndex > pres.Slides.Count Then Return False
            If toIndex < 1 OrElse toIndex > pres.Slides.Count Then Return False
            pres.Slides(fromIndex).MoveTo(toIndex)
            GlobalStatusStrip.ShowInfo($"已将幻灯片从第 {fromIndex} 移到第 {toIndex}")
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteMoveSlide 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 设置幻灯片版式
    ''' </summary>
    Private Function ExecuteSetSlideLayout(params As JToken, pres As Object) As Boolean
        Try
            Dim layoutName = If(params("layout")?.ToString(), "titleandcontent").ToLower()
            Dim slideIndex As Integer = If(params("slideIndex") IsNot Nothing, params("slideIndex").Value(Of Integer)(), -1)
            Dim targetSlide As Object
            If slideIndex = -1 OrElse slideIndex = 0 Then
                targetSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
            ElseIf slideIndex >= 1 AndAlso slideIndex <= pres.Slides.Count Then
                targetSlide = pres.Slides(slideIndex)
            Else
                Return False
            End If
            Dim layoutIndex As Integer = 2
            Select Case layoutName
                Case "title" : layoutIndex = 1
                Case "titleandcontent", "titleandbody" : layoutIndex = 2
                Case "blank" : layoutIndex = 12
                Case "titleonly" : layoutIndex = 11
            End Select
            targetSlide.Layout = layoutIndex
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteSetSlideLayout 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 应用主题
    ''' </summary>
    Private Function ExecuteApplyTheme(params As JToken, pres As Object) As Boolean
        Try
            Dim themeFile = If(params("themeFile")?.ToString(), "")
            Dim themeName = If(params("themeName")?.ToString(), "")
            If Not String.IsNullOrEmpty(themeFile) AndAlso IO.File.Exists(themeFile) Then
                pres.ApplyTheme(themeFile)
            ElseIf Not String.IsNullOrEmpty(themeName) Then
                GlobalStatusStrip.ShowInfo($"内置主题 '{themeName}' 请通过 WPS/PPT 设计菜单手动应用")
            End If
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteApplyTheme 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 添加演讲备注
    ''' </summary>
    Private Function ExecuteAddSpeakerNotes(params As JToken, pres As Object) As Boolean
        Try
            Dim notes = If(params("notes")?.ToString(), "")
            Dim slideIndex As Integer = If(params("slideIndex") IsNot Nothing, params("slideIndex").Value(Of Integer)(), -1)
            Dim targetSlide As Object
            If slideIndex = -1 OrElse slideIndex = 0 Then
                targetSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
            ElseIf slideIndex >= 1 AndAlso slideIndex <= pres.Slides.Count Then
                targetSlide = pres.Slides(slideIndex)
            Else
                Return False
            End If
            targetSlide.NotesPage.Shapes(2).TextFrame.TextRange.Text = notes
            GlobalStatusStrip.ShowInfo("已添加演讲备注")
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteAddSpeakerNotes 出错: {ex.Message}")
            Return False
        End Try
    End Function

#End Region

#Region "排版功能"

    ''' <summary>
    ''' 处理排版JSON响应（支持规则模式）
    ''' </summary>
    Protected Overrides Sub HandleApplyDocumentPlanItem(jsonDoc As JObject)
        Try
            ' 检测是否为规则模式
            If jsonDoc("rules") IsNot Nothing AndAlso jsonDoc("rules").Type = JTokenType.Array Then
                ApplyReformatRules(jsonDoc)
                Return
            End If

            ' 无效格式
            GlobalStatusStrip.ShowWarning("排版响应格式无效")

        Catch ex As Exception
            Debug.WriteLine("HandleApplyDocumentPlanItem 错误: " & ex.Message)
            GlobalStatusStrip.ShowWarning("排版应用出错: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 应用规则模式的排版
    ''' </summary>
    Private Sub ApplyReformatRules(jsonDoc As JObject)
        Try
            Dim rules = jsonDoc("rules").ToObject(Of List(Of JObject))()
            Dim sampleClassification = jsonDoc("sampleClassification")?.ToObject(Of List(Of JObject))()

            If rules Is Nothing OrElse rules.Count = 0 Then
                GlobalStatusStrip.ShowWarning("没有收到有效的排版规则")
                Return
            End If

            ' 构建规则字典
            Dim ruleDict As New Dictionary(Of String, JObject)()
            For Each rule In rules
                Dim ruleType = rule("type")?.ToString()
                If Not String.IsNullOrEmpty(ruleType) AndAlso rule("formatting") IsNot Nothing Then
                    ruleDict(ruleType) = DirectCast(rule("formatting"), JObject)
                End If
            Next

            ' 检查上下文
            If _reformatShapes Is Nothing OrElse _reformatShapes.Count = 0 Then
                GlobalStatusStrip.ShowWarning("排版上下文丢失，请重新选择内容并排版")
                Return
            End If

            ' 基于样本分类推断规则
            Dim sampleRuleMap As New Dictionary(Of Integer, String)()
            If sampleClassification IsNot Nothing Then
                For Each sc In sampleClassification
                    Dim idx = sc("sampleIndex")?.ToObject(Of Integer)()
                    Dim appliedRule = sc("appliedRule")?.ToString()
                    If idx IsNot Nothing AndAlso Not String.IsNullOrEmpty(appliedRule) Then
                        sampleRuleMap(idx) = appliedRule
                    End If
                Next
            End If

            ' 应用格式到所有形状
            Dim appliedCount As Integer = 0
            Dim defaultRule As String = If(ruleDict.ContainsKey("body"), "body", ruleDict.Keys.FirstOrDefault())

            ' 保存排版前快照（用于PPT撤销，因PPT不支持UndoRecord）
            _reformatUndoSnapshots = New List(Of ShapeFormatSnapshot)()
            For i As Integer = 0 To _reformatShapes.Count - 1
                Try
                    Dim shp = DirectCast(_reformatShapes(i), Microsoft.Office.Interop.PowerPoint.Shape)
                    _reformatUndoSnapshots.Add(ShapeFormatSnapshot.Capture(shp, i))
                Catch ex As Exception
                    Debug.WriteLine($"捕获形状{i}快照失败: {ex.Message}")
                End Try
            Next

            For i As Integer = 0 To _reformatShapes.Count - 1
                Try
                    Dim shp = DirectCast(_reformatShapes(i), Microsoft.Office.Interop.PowerPoint.Shape)
                    Dim shapeType = If(i < _reformatTypes.Count, _reformatTypes(i), "")

                    ' 确定使用哪个规则
                    Dim ruleToApply As String = defaultRule

                    If sampleRuleMap.ContainsKey(i) Then
                        ruleToApply = sampleRuleMap(i)
                    Else
                        ' 基于形状类型推断规则
                        If shapeType.Contains("标题") Then
                            For Each key In ruleDict.Keys
                                If key.ToLower().Contains("title") OrElse key = "标题" Then
                                    ruleToApply = key
                                    Exit For
                                End If
                            Next
                        ElseIf shapeType.Contains("副标题") Then
                            For Each key In ruleDict.Keys
                                If key.ToLower().Contains("subtitle") OrElse key = "副标题" Then
                                    ruleToApply = key
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    ' 应用规则
                    If ruleDict.ContainsKey(ruleToApply) Then
                        Dim formatting = ruleDict(ruleToApply)
                        ApplyFormattingToShape(shp, formatting)
                        appliedCount += 1
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"应用形状 {i} 格式失败: " & ex.Message)
                End Try
            Next

            ' 保留 _reformatShapes/_reformatTypes 用于撤销快照恢复
            ' 撤销或新排版开始时会重新设置

            GlobalStatusStrip.ShowInfo($"排版完成，共处理 {appliedCount} 个文本框")

        Catch ex As Exception
            Debug.WriteLine("ApplyReformatRules 错误: " & ex.Message)
            GlobalStatusStrip.ShowWarning("应用排版规则出错: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 应用格式化属性到形状
    ''' </summary>
    Private Sub ApplyFormattingToShape(shp As Microsoft.Office.Interop.PowerPoint.Shape, formatting As JObject)
        Try
            If shp.HasTextFrame <> Microsoft.Office.Core.MsoTriState.msoTrue Then Return

            Dim textRange = shp.TextFrame.TextRange

            ' 中文字体
            If formatting("fontNameCN") IsNot Nothing Then
                Try
                    textRange.Font.NameFarEast = formatting("fontNameCN").ToString()
                Catch
                End Try
            End If

            ' 英文字体
            If formatting("fontNameEN") IsNot Nothing Then
                Try
                    textRange.Font.Name = formatting("fontNameEN").ToString()
                Catch
                End Try
            End If

            ' 字号
            If formatting("fontSize") IsNot Nothing Then
                Dim fontSize As Single = 0
                Single.TryParse(formatting("fontSize").ToString(), fontSize)
                If fontSize > 0 Then
                    Try
                        textRange.Font.Size = fontSize
                    Catch
                    End Try
                End If
            End If

            ' 加粗
            If formatting("bold") IsNot Nothing Then
                Try
                    Dim bold As Boolean = formatting("bold").ToObject(Of Boolean)()
                    textRange.Font.Bold = If(bold, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse)
                Catch
                End Try
            End If

            ' 对齐方式
            If formatting("alignment") IsNot Nothing Then
                Dim alignment As String = formatting("alignment").ToString().ToLower()
                Try
                    Select Case alignment
                        Case "left"
                            textRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft
                        Case "center"
                            textRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter
                        Case "right"
                            textRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignRight
                        Case "justify"
                            textRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignJustify
                    End Select
                Catch
                End Try
            End If

        Catch ex As Exception
            Debug.WriteLine("ApplyFormattingToShape 出错: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 从快照恢复格式（PPT专用撤销方式）
    ''' </summary>
    Private Sub RestoreFormattingSnapshots()
        If _reformatUndoSnapshots Is Nothing OrElse _reformatUndoSnapshots.Count = 0 Then
            Debug.WriteLine("没有排版快照可恢复")
            Return
        End If

        Dim restoredCount As Integer = 0
        For Each snap In _reformatUndoSnapshots
            Try
                If snap.ShapeIndex >= 0 AndAlso snap.ShapeIndex < _reformatShapes.Count Then
                    Dim shp = DirectCast(_reformatShapes(snap.ShapeIndex), Microsoft.Office.Interop.PowerPoint.Shape)
                    snap.Restore(shp)
                    restoredCount += 1
                End If
            Catch ex As Exception
                Debug.WriteLine($"恢复形状{snap.ShapeIndex}失败: {ex.Message}")
            End Try
        Next
        Debug.WriteLine($"从快照恢复了 {restoredCount} 个形状的格式")

        ' 清理快照和上下文
        _reformatUndoSnapshots = Nothing
        _reformatShapes = Nothing
        _reformatTypes = Nothing
    End Sub

    ''' <summary>
    ''' 撤销排版（PPT重写：优先从快照恢复，失败后回退到基类Undo）
    ''' </summary>
    Protected Overrides Sub HandleUndoReformat()
        ' 优先尝试从快照恢复（最可靠）
        If _reformatUndoSnapshots IsNot Nothing AndAlso _reformatUndoSnapshots.Count > 0 Then
            Try
                RestoreFormattingSnapshots()
                GlobalStatusStrip.ShowInfo("已撤销排版操作")
                Return
            Catch ex As Exception
                Debug.WriteLine($"快照恢复失败，回退到Undo: {ex.Message}")
            End Try
        End If

        ' 回退到基类Undo机制
        MyBase.HandleUndoReformat()
    End Sub

#End Region

End Class

