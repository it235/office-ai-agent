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
Imports System.Text.JSON
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.UI.WebControls
Imports System.Windows.Forms
Imports System.Windows.Forms.ListBox
Imports Markdig
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon
Public Class ChatControl
    Inherits BaseChatControl


    Private sheetContentItems As New Dictionary(Of String, Tuple(Of System.Windows.Forms.Label, System.Windows.Forms.Button))

    ' 排版上下文：存储待格式化的形状和类型信息
    Private _reformatShapes As List(Of Object) = Nothing
    Private _reformatTypes As List(Of String) = Nothing

    ''' <summary>
    ''' 设置排版上下文，用于规则匹配后应用格式
    ''' </summary>
    Public Sub SetReformatContext(shapes As List(Of Object), types As List(Of String))
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

    Private Function GetSelectionDetails(selection As Object) As String
        Try
            Dim details As New StringBuilder()
            Dim ppSelection = TryCast(selection, Microsoft.Office.Interop.PowerPoint.Selection)

            If ppSelection Is Nothing Then
                Return "未选中任何内容"
            End If

            ' 添加基本信息
            details.AppendLine($"选择类型: {ppSelection.Type}")

            If ppSelection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes Then
                Dim shapeRange = ppSelection.ShapeRange
                details.AppendLine($"形状数量: {shapeRange.Count}")
                For i = 1 To shapeRange.Count
                    details.AppendLine($"形状 {i} 类型: {shapeRange(i).Type}")
                    ' 检查是否是表格
                    If shapeRange(i).HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                        Dim table = shapeRange(i).Table
                        details.AppendLine($"表格大小: {table.Rows.Count}行 x {table.Columns.Count}列")
                    ElseIf shapeRange(i).HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                        details.AppendLine($"形状 {i} 文本长度: {shapeRange(i).TextFrame.TextRange.Length}")
                    End If
                Next

            ElseIf ppSelection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText Then
                Dim textRange = ppSelection.TextRange
                details.AppendLine($"文本长度: {textRange.Length}")
                details.AppendLine($"字符数: {textRange.Length}")

            ElseIf ppSelection.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionSlides Then
                Dim slideRange = ppSelection.SlideRange
                details.AppendLine($"选中幻灯片数: {slideRange.Count}")
                For i = 1 To slideRange.Count
                    details.AppendLine($"幻灯片 {i} 标题: {slideRange(i).Name}")
                Next
            End If

            Return details.ToString()
        Catch ex As Exception
            Return $"获取选择详情时出错: {ex.Message}"
        End Try
    End Function

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
            ' 创建一个新的PowerPoint应用程序实例
            Dim pptApp As New Microsoft.Office.Interop.PowerPoint.Application()
            pptApp.Visible = Microsoft.Office.Core.MsoTriState.msoFalse

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
                    Dim slide As Microsoft.Office.Interop.PowerPoint.Slide = presentation.Slides(slideIndex)
                    contentBuilder.AppendLine($"=== 幻灯片 {slideIndex} ===")

                    ' 遍历幻灯片中的形状
                    For Each shape As Microsoft.Office.Interop.PowerPoint.Shape In slide.Shapes
                        Try
                            ' 检查是否有文本框架
                            If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                If shape.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                    Dim text As String = shape.TextFrame.TextRange.Text.Trim()
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
                                End If
                            End If

                            ' 检查是否是表格
                            If shape.HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                Dim table = shape.Table
                                contentBuilder.AppendLine($"  [表格 {table.Rows.Count}行×{table.Columns.Count}列]")
                                ' 读取表格内容（限制行数）
                                Dim maxRows = Math.Min(table.Rows.Count, 10)
                                For rowIdx = 1 To maxRows
                                    Dim rowContent As New StringBuilder("    ")
                                    For colIdx = 1 To table.Columns.Count
                                        Try
                                            Dim cellText = table.Cell(rowIdx, colIdx).Shape.TextFrame.TextRange.Text.Trim()
                                            If cellText.Length > 20 Then cellText = cellText.Substring(0, 17) & "..."
                                            rowContent.Append(cellText & " | ")
                                        Catch
                                        End Try
                                    Next
                                    contentBuilder.AppendLine(rowContent.ToString().TrimEnd(" |".ToCharArray()))
                                Next
                            End If
                        Catch shapeEx As Exception
                            Debug.WriteLine($"处理形状时出错: {shapeEx.Message}")
                        End Try
                    Next

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
                End If
                pptApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pptApp)
                GC.Collect()
                GC.WaitForPendingFinalizers()
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
                        Dim slide = slideRange(i)
                        contentBuilder.AppendLine($"幻灯片 {slide.SlideIndex}:")

                        ' 获取幻灯片标题
                        Dim title As String = ""
                        For Each shape In slide.Shapes
                            If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder Then
                                If shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle Then
                                    If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                        title = shape.TextFrame.TextRange.Text.Trim()
                                        Exit For
                                    End If
                                End If
                            End If
                        Next

                        If title <> "" Then
                            contentBuilder.AppendLine($"  标题: {title}")
                        Else
                            contentBuilder.AppendLine("  [无标题]")
                        End If

                        ' 获取幻灯片上的内容
                        Dim textShapesCount = 0

                        For Each shape In slide.Shapes
                            ' 跳过标题形状
                            If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder AndAlso
                           shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle Then
                                Continue For
                            End If

                            ' 处理文本形状
                            If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue AndAlso
                           shape.TextFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then

                                textShapesCount += 1
                                If textShapesCount > 3 Then Continue For ' 每张幻灯片最多处理3个文本框

                                Dim text = shape.TextFrame.TextRange.Text.Trim()
                                If text.Length > 0 Then
                                    ' 限制文本长度
                                    If text.Length > 200 Then
                                        contentBuilder.AppendLine($"  文本: {text.Substring(0, 200)}...")
                                    Else
                                        contentBuilder.AppendLine($"  文本: {text}")
                                    End If
                                End If
                            ElseIf shape.HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                contentBuilder.AppendLine("  [包含表格]")
                            ElseIf shape.Type = Microsoft.Office.Core.MsoShapeType.msoPicture Then
                                contentBuilder.AppendLine("  [包含图片]")
                            End If
                        Next

                        contentBuilder.AppendLine("  ---")
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
    Private Sub ProcessShapeSelection(builder As StringBuilder, selection As Microsoft.Office.Interop.PowerPoint.Selection)
        Try
            Dim shapeRange = selection.ShapeRange
            builder.AppendLine($"形状数量: {shapeRange.Count}")

            ' 遍历选中的形状
            For i = 1 To shapeRange.Count
                builder.AppendLine($"形状 {i}:")

                ' 检查是否是表格
                If shapeRange(i).HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    ProcessTable(builder, shapeRange(i).Table)
                ElseIf shapeRange(i).HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    ' 处理包含文本的形状
                    Dim textFrame = shapeRange(i).TextFrame
                    If textFrame.HasText = Microsoft.Office.Core.MsoTriState.msoTrue Then
                        Dim text = textFrame.TextRange.Text.Trim()
                        ' 限制文本长度
                        If text.Length > 1000 Then
                            builder.AppendLine(text.Substring(0, 1000) & "...")
                            builder.AppendLine($"[文本太长，仅显示前1000个字符，总计: {text.Length}个字符]")
                        Else
                            builder.AppendLine(text)
                        End If
                    Else
                        builder.AppendLine("[空文本框]")
                    End If
                ElseIf shapeRange(i).Type = Microsoft.Office.Core.MsoShapeType.msoPicture Then
                    ' 处理图片
                    builder.AppendLine("[图片]")
                    ' 尝试获取图片的替代文本（如果有）
                    If shapeRange(i).AlternativeText <> "" Then
                        builder.AppendLine($"替代文本: {shapeRange(i).AlternativeText}")
                    End If
                ElseIf shapeRange(i).Type = Microsoft.Office.Core.MsoShapeType.msoChart Then
                    ' 处理图表
                    builder.AppendLine("[图表]")
                    If shapeRange(i).AlternativeText <> "" Then
                        builder.AppendLine($"图表说明: {shapeRange(i).AlternativeText}")
                    End If
                ElseIf shapeRange(i).Type = Microsoft.Office.Core.MsoShapeType.msoSmartArt Then
                    ' 处理SmartArt
                    builder.AppendLine("[SmartArt图形]")
                Else
                    ' 其他类型的形状
                    builder.AppendLine($"[形状类型: {shapeRange(i).Type}]")
                End If

                ' 形状之间添加分隔线
                builder.AppendLine("---")
            Next

        Catch ex As Exception
            builder.AppendLine($"[处理形状时出错: {ex.Message}]")
        End Try
    End Sub

    ' 处理表格内容
    Private Sub ProcessTable(builder As StringBuilder, table As Microsoft.Office.Interop.PowerPoint.Table)
        Try
            builder.AppendLine($"表格: {table.Rows.Count}行 × {table.Columns.Count}列")

            ' 限制显示的行列数
            Dim maxRows As Integer = Math.Min(table.Rows.Count, 20)
            Dim maxCols As Integer = Math.Min(table.Columns.Count, 10)

            ' 处理表格头部（表格第一行）
            If table.Rows.Count > 0 Then
                ' 构建表头和分隔线
                Dim headerBuilder As New StringBuilder()
                Dim separatorBuilder As New StringBuilder()

                For col As Integer = 1 To maxCols
                    Try
                        Dim cellText As String = table.Cell(1, col).Shape.TextFrame.TextRange.Text.Trim()

                        ' 限制单元格文本长度
                        If cellText.Length > 20 Then
                            cellText = cellText.Substring(0, 17) & "..."
                        End If

                        ' 填充表头
                        If col > 1 Then
                            headerBuilder.Append(" | ")
                            separatorBuilder.Append("-+-")
                        End If
                        headerBuilder.Append(cellText)
                        separatorBuilder.Append(New String("-"c, Math.Max(cellText.Length, 3)))
                    Catch ex As Exception
                        ' 忽略单元格处理错误
                        If col > 1 Then
                            headerBuilder.Append(" | ")
                            separatorBuilder.Append("-+-")
                        End If
                        headerBuilder.Append("N/A")
                        separatorBuilder.Append("---")
                    End Try
                Next

                ' 添加表头和分隔线
                builder.AppendLine(headerBuilder.ToString())
                builder.AppendLine(separatorBuilder.ToString())
            End If

            ' 处理表格数据行
            For row As Integer = 2 To maxRows ' 从第2行开始（跳过表头）
                Dim rowBuilder As New StringBuilder()

                For col As Integer = 1 To maxCols
                    Try
                        Dim cellText As String = table.Cell(row, col).Shape.TextFrame.TextRange.Text.Trim()

                        ' 限制单元格文本长度
                        If cellText.Length > 20 Then
                            cellText = cellText.Substring(0, 17) & "..."
                        End If

                        ' 填充行数据
                        If col > 1 Then
                            rowBuilder.Append(" | ")
                        End If
                        rowBuilder.Append(cellText)
                    Catch ex As Exception
                        ' 忽略单元格处理错误
                        If col > 1 Then
                            rowBuilder.Append(" | ")
                        End If
                        rowBuilder.Append("N/A")
                    End Try
                Next

                ' 添加行数据
                builder.AppendLine(rowBuilder.ToString())
            Next

            ' 如果有更多行未显示，添加提示
            If table.Rows.Count > maxRows Then
                builder.AppendLine($"... [表格共有 {table.Rows.Count} 行，仅显示前 {maxRows} 行]")
            End If

            ' 如果有更多列未显示，添加提示
            If table.Columns.Count > maxCols Then
                builder.AppendLine($"... [表格共有 {table.Columns.Count} 列，仅显示前 {maxCols} 列]")
            End If

        Catch ex As Exception
            builder.AppendLine($"[处理表格内容时出错: {ex.Message}]")
        End Try
    End Sub

    ' 处理文本选择
    Private Sub ProcessTextSelection(builder As StringBuilder, selection As Microsoft.Office.Interop.PowerPoint.Selection)
        Try
            Dim textRange = selection.TextRange

            If textRange IsNot Nothing Then
                builder.AppendLine($"文本长度: {textRange.Length} 个字符")

                ' 获取文本内容并限制长度
                Dim text = textRange.Text.Trim()
                Dim maxLength As Integer = 2000

                If text.Length > maxLength Then
                    builder.AppendLine(text.Substring(0, maxLength) & "...")
                    builder.AppendLine($"[文本太长，仅显示前{maxLength}个字符，总计: {text.Length}个字符]")
                Else
                    builder.AppendLine(text)
                End If
            Else
                builder.AppendLine("[无法获取文本内容]")
            End If

        Catch ex As Exception
            builder.AppendLine($"[处理文本选择时出错: {ex.Message}]")
        End Try
    End Sub

    ' 处理幻灯片选择
    Private Sub ProcessSlideSelection(builder As StringBuilder, selection As Microsoft.Office.Interop.PowerPoint.Selection)
        Try
            Dim slideRange = selection.SlideRange
            builder.AppendLine($"选中幻灯片数: {slideRange.Count}")

            ' 限制处理的幻灯片数量
            Dim maxSlides As Integer = Math.Min(slideRange.Count, 10)

            For i = 1 To maxSlides
                Dim slide = slideRange(i)
                builder.AppendLine($"幻灯片 {slide.SlideIndex}:")

                ' 获取幻灯片标题
                Dim title As String = GetSlideTitle(slide)
                If Not String.IsNullOrEmpty(title) Then
                    builder.AppendLine($"标题: {title}")
                End If

                ' 获取幻灯片上的内容
                builder.AppendLine("内容:")
                Dim slideContent = GetSlideContent(slide)
                builder.AppendLine(slideContent)

                ' 添加分隔线
                builder.AppendLine("---")
            Next

            ' 如果有更多幻灯片未显示，添加提示
            If slideRange.Count > maxSlides Then
                builder.AppendLine($"... [共选中 {slideRange.Count} 张幻灯片，仅显示前 {maxSlides} 张]")
            End If

        Catch ex As Exception
            builder.AppendLine($"[处理幻灯片选择时出错: {ex.Message}]")
        End Try
    End Sub

    ' 获取幻灯片标题
    Private Function GetSlideTitle(slide As Microsoft.Office.Interop.PowerPoint.Slide) As String
        Try
            ' 检查幻灯片是否有标题占位符
            For Each shape In slide.Shapes
                If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder Then
                    If shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle Then
                        If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            Return shape.TextFrame.TextRange.Text.Trim()
                        End If
                    End If
                End If
            Next

            ' 如果没有找到标题占位符，尝试查找任何可能的标题
            For Each shape In slide.Shapes
                If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    Dim text = shape.TextFrame.TextRange.Text.Trim()
                    If Not String.IsNullOrEmpty(text) AndAlso text.Length < 100 Then
                        Return text ' 假设第一个简短文本是标题
                    End If
                End If
            Next

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
            For Each shape In slide.Shapes
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

                    Dim text = shape.TextFrame.TextRange.Text.Trim()
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
            Next

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
            ' 获取当前活动工作簿的路径
            If Globals.ThisAddIn.Application.ActiveWorkbook IsNot Nothing Then
                Return Globals.ThisAddIn.Application.ActiveWorkbook.Path
            End If
        Catch ex As Exception
            Debug.WriteLine($"获取当前工作目录时出错: {ex.Message}")
        End Try

        ' 如果无法获取工作簿路径，则返回应用程序目录
        Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
    End Function


    Protected Overrides Sub CheckAndCompleteProcessingHook(_finalUuid As String, allPlainMarkdownBuffer As StringBuilder)
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
                ExecuteJavaScriptAsyncJS($"removeContinuationActions('{uuid}');")
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
                ExecuteJavaScriptAsyncJS($"removeTemplateActions('{uuid}');")
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
                    Dim slideIndex = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex
                    snapshot("currentSlide") = slideIndex
                Catch
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
    Protected Overrides Function ExecuteJsonCommand(jsonCode As String, preview As Boolean) As Boolean
        Try
            ' 预览模式下跳过自动执行（排版/校对模式的JSON用于预览，由用户手动点击应用）
            If IsInPreviewMode() Then
                Debug.WriteLine($"[PPTChatControl] 预览模式({GetCurrentResponseMode()})下跳过JSON命令自动执行")
                Return True ' 返回True表示"成功处理"，避免显示错误
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
                    Return True
                End If
            End If

            ' 执行所有命令
            Dim successCount = 0
            Dim failCount = 0

            For Each cmd In commands
                If cmd.Type = JTokenType.Object Then
                    Dim cmdObj = CType(cmd, JObject)
                    If ExecutePPTCommand(cmdObj) Then
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
                    Return True
                End If
            End If

            ' 执行命令
            Dim success = ExecutePPTCommand(commandJson)

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
    Private Function ExecutePPTCommand(commandJson As JObject) As Boolean
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
                    Return ExecuteCreateSlides(params, pres)
                Case "addanimation"
                    Return ExecuteAddAnimation(params, pres)
                Case "applytransition"
                    Return ExecuteApplyTransition(params, pres)
                Case "beautifyslides"
                    Return ExecuteBeautifySlides(params, pres)
                Case Else
                    Debug.WriteLine($"不支持的PPT命令: {command}")
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
            Dim slide = pres.Slides.Add(slideIndex, 11)

            ' 设置标题
            If Not String.IsNullOrEmpty(title) Then
                For Each shape In slide.Shapes
                    If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder Then
                        If shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle Then
                            shape.TextFrame.TextRange.Text = title
                            Exit For
                        End If
                    End If
                Next
            End If

            ' 如果有内容，添加文本框
            If Not String.IsNullOrEmpty(content) Then
                Dim textBox = slide.Shapes.AddTextbox(
                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                    50, 150, 600, 300)
                textBox.TextFrame.TextRange.Text = content
            End If

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
                Dim slide = pres.Slides.Add(startIndex + i, layoutType)

                ' 填充标题
                If Not String.IsNullOrEmpty(title) Then
                    For Each shape In slide.Shapes
                        If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder Then
                            If shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle Then
                                shape.TextFrame.TextRange.Text = title
                                Exit For
                            End If
                        End If
                    Next
                End If

                ' 填充内容
                If Not String.IsNullOrEmpty(content) Then
                    Dim contentFilled = False
                    For Each shape In slide.Shapes
                        If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder Then
                            If shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderBody OrElse
                               shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle Then
                                shape.TextFrame.TextRange.Text = content
                                contentFilled = True
                                Exit For
                            End If
                        End If
                    Next

                    ' 如果没有内容占位符，添加文本框
                    If Not contentFilled Then
                        Dim textBox = slide.Shapes.AddTextbox(
                            Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                            50, 150, 620, 350)
                        textBox.TextFrame.TextRange.Text = content
                        textBox.TextFrame.TextRange.Font.Size = 18
                    End If
                End If
            Next

            ShareRibbon.GlobalStatusStrip.ShowInfo($"成功创建 {slidesArray.Count} 张幻灯片")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteCreateSlides 出错: {ex.Message}")
            Return False
        End Try
    End Function

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

            Dim slide As Object
            If slideIndex < 0 Then
                slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide
            Else
                slide = pres.Slides(Math.Min(slideIndex + 1, pres.Slides.Count))
            End If

            Dim timeline = slide.TimeLine
            Dim msoEffect = GetMsoAnimEffect(effect)

            For Each shape In slide.Shapes
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

                If shouldAnimate AndAlso shape.HasTextFrame Then
                    Try
                        timeline.MainSequence.AddEffect(shape, msoEffect)
                    Catch
                        ' 某些形状可能不支持动画
                    End Try
                End If
            Next

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

            Dim slidesToProcess As New List(Of Object)()
            If scope.ToLower() = "all" Then
                For Each slide In pres.Slides
                    slidesToProcess.Add(slide)
                Next
            Else
                slidesToProcess.Add(Globals.ThisAddIn.Application.ActiveWindow.View.Slide)
            End If

            For Each slide In slidesToProcess
                slide.SlideShowTransition.EntryEffect = transEffect
                slide.SlideShowTransition.Duration = duration
                slide.SlideShowTransition.AdvanceOnClick = True
            Next

            ShareRibbon.GlobalStatusStrip.ShowInfo($"已为 {slidesToProcess.Count} 张幻灯片应用切换效果")
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

            Dim slidesToProcess As New List(Of Object)()
            If scope.ToLower() = "all" Then
                For Each slide In pres.Slides
                    slidesToProcess.Add(slide)
                Next
            Else
                slidesToProcess.Add(Globals.ThisAddIn.Application.ActiveWindow.View.Slide)
            End If

            For Each slide In slidesToProcess
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
                For Each shape In slide.Shapes
                    If shape.HasTextFrame Then
                        Dim isTitle = False
                        If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPlaceholder Then
                            isTitle = (shape.PlaceholderFormat.Type = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle)
                        End If

                        Dim fontTheme = If(isTitle, theme?("titleFont"), theme?("bodyFont"))
                        If fontTheme IsNot Nothing Then
                            ApplyFontStyle(shape.TextFrame.TextRange, fontTheme)
                        End If
                    End If
                Next
            Next

            ShareRibbon.GlobalStatusStrip.ShowInfo($"已美化 {slidesToProcess.Count} 张幻灯片")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteBeautifySlides 出错: {ex.Message}")
            Return False
        End Try
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

            ' 清理上下文
            _reformatShapes = Nothing
            _reformatTypes = Nothing

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

#End Region

End Class

