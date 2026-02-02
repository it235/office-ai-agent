' PowerPointAi\PowerPointJsonCommandSchema.vb
' PowerPoint JSON命令Schema定义和校验

Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json.Schema

''' <summary>
''' PowerPoint JSON命令Schema和校验器
''' </summary>
Public Class PowerPointJsonCommandSchema

    ''' <summary>
    ''' 支持的命令类型 (22个命令覆盖主流PowerPoint操作场景)
    ''' 幻灯片操作: InsertSlide, DeleteSlide, DuplicateSlide, MoveSlide, CreateSlides
    ''' 内容操作: InsertText, FormatText, InsertShape, InsertImage, InsertTable
    ''' 样式和动画: FormatSlide, AddAnimation, ApplyTransition, BeautifySlides, SetSlideLayout
    ''' 高级功能: InsertChart, InsertVideo, AddSpeakerNotes, SetSlideShow
    ''' 母版和主题: ApplyTheme, EditSlideMaster
    ''' VBA回退: ExecuteVBA
    ''' </summary>
    Public Shared ReadOnly SupportedCommands As String() = {
        "InsertSlide",
        "DeleteSlide",
        "DuplicateSlide",
        "MoveSlide",
        "CreateSlides",
        "InsertText",
        "FormatText",
        "InsertShape",
        "InsertImage",
        "InsertTable",
        "FormatSlide",
        "AddAnimation",
        "ApplyTransition",
        "BeautifySlides",
        "SetSlideLayout",
        "InsertChart",
        "InsertVideo",
        "AddSpeakerNotes",
        "SetSlideShow",
        "ApplyTheme",
        "EditSlideMaster",
        "ExecuteVBA"
    }

    ''' <summary>
    ''' 获取严格的JSON Schema定义（用于约束大模型输出）
    ''' </summary>
    Public Shared Function GetStrictJsonSchemaPrompt() As String
        Return "
【重要】你必须且只能返回以下两种JSON格式之一：

格式1 - 单个命令：
```json
{
  ""command"": ""InsertSlide"",
  ""params"": {
    ""position"": ""end"",
    ""layout"": ""title"",
    ""title"": ""幻灯片标题""
  }
}
```

格式2 - 多个命令（批量操作）：
```json
{
  ""commands"": [
    {""command"": ""InsertSlide"", ""params"": {""title"": ""第一页""}},
    {""command"": ""AddAnimation"", ""params"": {""effect"": ""fadeIn"", ""scope"": ""all""}}
  ]
}
```

【绝对禁止的格式】
- 禁止 {""command"": ""xxx"", ""actions"": [...]}
- 禁止 {""command"": ""xxx"", ""title"": ""...""} (缺少params包装)
- 禁止 {""operations"": [...]}

【PowerPoint支持的22个命令】

=== 幻灯片操作 (5个) ===
1. InsertSlide - 插入幻灯片 {position:current/end, layout:title/titleAndContent/blank, title:可选, content:可选}
2. DeleteSlide - 删除幻灯片 {slideIndex:必需,-1表示当前}
3. DuplicateSlide - 复制幻灯片 {slideIndex:必需, insertAfter:可选}
4. MoveSlide - 移动幻灯片 {fromIndex:必需, toIndex:必需}
5. CreateSlides - 批量创建 {slides:数组,每个含title/content/layout}

=== 内容操作 (5个) ===
6. InsertText - 插入文本 {content:必需, slideIndex:-1表示当前, x/y:可选位置}
7. FormatText - 格式化文本 {slideIndex:可选, shapeIndex:可选, bold/italic/fontSize/fontName/color}
8. InsertShape - 插入形状 {shapeType:rectangle/oval/arrow等, x:必需, y:必需, width/height:可选}
9. InsertImage - 插入图片 {imagePath:必需, slideIndex:可选, x/y/width/height:可选}
10. InsertTable - 插入表格 {rows:必需, cols:必需, data:可选, slideIndex:可选}

=== 样式和动画 (5个) ===
11. FormatSlide - 格式化幻灯片 {slideIndex:可选, background:颜色/图片路径, layout:可选}
12. AddAnimation - 添加动画 {effect:fadeIn/flyIn/zoom/wipe/appear, slideIndex:可选, targetShapes:all/title/content}
13. ApplyTransition - 切换效果 {transitionType:fade/push/wipe/split, scope:all/current, duration:秒}
14. BeautifySlides - 美化幻灯片 {scope:all/current, theme:{background/titleFont/bodyFont}}
15. SetSlideLayout - 设置布局 {slideIndex:可选, layout:title/titleAndContent/twoContent/blank/comparison}

=== 高级功能 (4个) ===
16. InsertChart - 插入图表 {chartType:column/line/pie/bar, data:二维数组, slideIndex:可选, title:可选}
17. InsertVideo - 插入视频 {videoPath:必需, slideIndex:可选, x/y/width/height:可选, autoPlay:可选}
18. AddSpeakerNotes - 演讲备注 {slideIndex:可选, notes:必需}
19. SetSlideShow - 放映设置 {loopUntilEsc:可选, showWithNarration:可选, advanceMode:manual/automatic}

=== 母版和主题 (2个) ===
20. ApplyTheme - 应用主题 {themeName:可选内置主题名, themeFile:可选主题文件路径}
21. EditSlideMaster - 编辑母版 {background:可选, titleFont:可选, bodyFont:可选}

=== VBA回退 (1个) ===
22. ExecuteVBA - 执行VBA代码 {code:必需,完整的Sub或Function代码}
    当以上命令无法满足需求时,生成VBA代码作为回退方案

【slideIndex说明】
- -1 或不填表示当前幻灯片
- 0 表示第一张幻灯片
- 正数表示具体幻灯片索引

【重要决策规则】
1. 优先使用上述22个命令处理用户需求
2. 复杂需求无法用命令实现时，使用ExecuteVBA生成VBA代码
3. 翻译需求请告知用户使用工具栏的""翻译""按钮
4. 如果用户需求不明确，直接用中文询问"
    End Function

    ''' <summary>
    ''' 获取格式校验失败的重试提示（Self-check机制）
    ''' </summary>
    Public Shared Function GetFormatCorrectionPrompt(originalJson As String, errorMessage As String) As String
        Return $"你之前返回的JSON格式不符合规范:

【错误原因】{errorMessage}

【你返回的内容】
{originalJson}

【正确格式示例】
单命令:
{{""command"": ""InsertSlide"", ""params"": {{""position"": ""end"", ""title"": ""标题""}}}}

多命令:
{{""commands"": [{{""command"": ""InsertSlide"", ""params"": {{""title"": ""第一页""}}}}, {{""command"": ""InsertText"", ""params"": {{""content"": ""内容""}}}}]}}

请严格按照上述格式重新返回JSON命令。"
    End Function

    ''' <summary>
    ''' 验证整个JSON响应结构是否符合规范
    ''' </summary>
    Public Shared Function ValidateJsonStructure(jsonText As String, ByRef errorMessage As String, ByRef normalizedJson As JToken) As Boolean
        Try
            errorMessage = ""
            normalizedJson = Nothing

            Dim token = JToken.Parse(jsonText)
            If token.Type <> JTokenType.Object Then
                errorMessage = "响应必须是JSON对象"
                Return False
            End If

            Dim jsonObj = CType(token, JObject)

            ' 检查是否是 commands 数组格式
            If jsonObj("commands") IsNot Nothing Then
                If jsonObj("commands").Type <> JTokenType.Array Then
                    errorMessage = "commands必须是数组"
                    Return False
                End If

                ' 验证数组中的每个命令
                Dim commands = CType(jsonObj("commands"), JArray)
                For i As Integer = 0 To commands.Count - 1
                    Dim cmd = commands(i)
                    If cmd.Type <> JTokenType.Object Then
                        errorMessage = $"commands[{i}]必须是对象"
                        Return False
                    End If
                    
                    ' 标准化并验证每个命令
                    Dim cmdObj = CType(cmd, JObject)
                    cmdObj = NormalizeCommandStructure(cmdObj)
                    commands(i) = cmdObj
                    
                    Dim cmdError As String = ""
                    If Not ValidateCommand(cmdObj, cmdError) Then
                        errorMessage = $"commands[{i}]: {cmdError}"
                        Return False
                    End If
                Next
                
                normalizedJson = jsonObj
                Return True
            End If

            ' 检查是否有禁止的格式
            If jsonObj("actions") IsNot Nothing Then
                errorMessage = "禁止使用actions格式，请使用commands数组"
                Return False
            End If

            If jsonObj("operations") IsNot Nothing Then
                errorMessage = "禁止使用operations格式，请使用commands数组"
                Return False
            End If

            ' 单命令格式
            If jsonObj("command") IsNot Nothing Then
                jsonObj = NormalizeCommandStructure(jsonObj)
                Dim cmdError As String = ""
                If Not ValidateCommand(jsonObj, cmdError) Then
                    errorMessage = cmdError
                    Return False
                End If
                normalizedJson = jsonObj
                Return True
            End If

            errorMessage = "缺少command或commands字段"
            Return False

        Catch ex As Newtonsoft.Json.JsonReaderException
            errorMessage = $"JSON解析失败: {ex.Message}"
            Return False
        Catch ex As Exception
            errorMessage = $"验证异常: {ex.Message}"
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 标准化JSON命令结构 - 将扁平结构自动包装到params中
    ''' </summary>
    Public Shared Function NormalizeCommandStructure(json As JObject) As JObject
        Try
            ' 检查是否已有params字段
            If json("params") IsNot Nothing Then
                Return json
            End If

            ' 检查是否有command字段
            Dim command = json("command")?.ToString()
            If String.IsNullOrEmpty(command) Then
                Return json
            End If

            ' 需要移到params中的字段
            Dim topLevelFields As String() = {"command"}
            Dim paramsFields As New JObject()

            For Each prop In json.Properties().ToList()
                If Not topLevelFields.Contains(prop.Name, StringComparer.OrdinalIgnoreCase) Then
                    paramsFields(prop.Name) = prop.Value
                    json.Remove(prop.Name)
                End If
            Next

            If paramsFields.Count > 0 Then
                json("params") = paramsFields
            End If

            Return json
        Catch ex As Exception
            Debug.WriteLine($"NormalizeCommandStructure 出错: {ex.Message}")
            Return json
        End Try
    End Function

    ''' <summary>
    ''' 校验JSON命令是否有效
    ''' </summary>
    Public Shared Function ValidateCommand(json As JObject, ByRef errorMessage As String) As Boolean
        Try
            errorMessage = ""
            
            ' 首先进行结构标准化
            json = NormalizeCommandStructure(json)
            
            Dim command = json("command")?.ToString()
            If String.IsNullOrEmpty(command) Then
                errorMessage = "缺少command字段"
                Return False
            End If
            
            If Not SupportedCommands.Any(Function(c) c.Equals(command, StringComparison.OrdinalIgnoreCase)) Then
                errorMessage = $"不支持的命令: {command}。支持的命令: {String.Join(", ", SupportedCommands)}"
                Return False
            End If
            
            Dim params = json("params")
            If params Is Nothing Then
                errorMessage = "缺少params字段"
                Return False
            End If
            
            ' 根据命令类型校验参数
            Select Case command.ToLower()
                ' === 幻灯片操作 ===
                Case "insertslide"
                    Return ValidateInsertSlide(params, errorMessage)
                Case "deleteslide"
                    Return ValidateDeleteSlide(params, errorMessage)
                Case "duplicateslide"
                    Return ValidateDuplicateSlide(params, errorMessage)
                Case "moveslide"
                    Return ValidateMoveSlide(params, errorMessage)
                Case "createslides"
                    Return ValidateCreateSlides(params, errorMessage)
                ' === 内容操作 ===
                Case "inserttext"
                    Return ValidateInsertText(params, errorMessage)
                Case "formattext"
                    Return ValidateFormatText(params, errorMessage)
                Case "insertshape"
                    Return ValidateInsertShape(params, errorMessage)
                Case "insertimage"
                    Return ValidateInsertImage(params, errorMessage)
                Case "inserttable"
                    Return ValidateInsertTable(params, errorMessage)
                ' === 样式和动画 ===
                Case "formatslide"
                    Return ValidateFormatSlide(params, errorMessage)
                Case "addanimation"
                    Return ValidateAddAnimation(params, errorMessage)
                Case "applytransition"
                    Return ValidateApplyTransition(params, errorMessage)
                Case "beautifyslides"
                    Return ValidateBeautifySlides(params, errorMessage)
                Case "setslidelayout"
                    Return ValidateSetSlideLayout(params, errorMessage)
                ' === 高级功能 ===
                Case "insertchart"
                    Return ValidateInsertChart(params, errorMessage)
                Case "insertvideo"
                    Return ValidateInsertVideo(params, errorMessage)
                Case "addspeakernotes"
                    Return ValidateAddSpeakerNotes(params, errorMessage)
                Case "setslideshow"
                    Return ValidateSetSlideShow(params, errorMessage)
                ' === 母版和主题 ===
                Case "applytheme"
                    Return ValidateApplyTheme(params, errorMessage)
                Case "editslidemaster"
                    Return ValidateEditSlideMaster(params, errorMessage)
                ' === VBA回退 ===
                Case "executevba"
                    Return ValidateExecuteVBA(params, errorMessage)
                Case Else
                    Return True
            End Select
            
        Catch ex As Exception
            errorMessage = $"JSON校验异常: {ex.Message}"
            Return False
        End Try
    End Function

    Private Shared Function ValidateInsertSlide(params As JToken, ByRef errorMessage As String) As Boolean
        ' InsertSlide参数都是可选的，基本验证通过
        Return True
    End Function

    Private Shared Function ValidateInsertText(params As JToken, ByRef errorMessage As String) As Boolean
        Dim content = params("content")?.ToString()
        ' 兼容处理：大模型在GENERAL_QUERY模式下可能返回text而不是content
        If String.IsNullOrEmpty(content) Then
            content = params("text")?.ToString()
            ' 如果text存在，将其复制到content字段以便后续执行
            If Not String.IsNullOrEmpty(content) AndAlso params.Type = JTokenType.Object Then
                CType(params, JObject)("content") = content
            End If
        End If
        If String.IsNullOrEmpty(content) Then
            errorMessage = "InsertText缺少content参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateInsertShape(params As JToken, ByRef errorMessage As String) As Boolean
        Dim shapeType = params("shapeType")?.ToString()
        If String.IsNullOrEmpty(shapeType) Then
            errorMessage = "InsertShape缺少shapeType参数"
            Return False
        End If
        
        If params("x") Is Nothing OrElse params("y") Is Nothing Then
            errorMessage = "InsertShape缺少x或y参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateFormatSlide(params As JToken, ByRef errorMessage As String) As Boolean
        ' FormatSlide至少需要一个格式化属性
        If params("background") Is Nothing AndAlso params("transition") Is Nothing AndAlso
           params("layout") Is Nothing Then
            errorMessage = "FormatSlide至少需要一个格式化属性(background/transition/layout)"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateInsertTable(params As JToken, ByRef errorMessage As String) As Boolean
        Dim rows = params("rows")
        Dim cols = params("cols")
        
        If rows Is Nothing OrElse cols Is Nothing Then
            errorMessage = "InsertTable缺少rows或cols参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateCreateSlides(params As JToken, ByRef errorMessage As String) As Boolean
        Dim slides = params("slides")
        If slides Is Nothing OrElse slides.Type <> JTokenType.Array Then
            errorMessage = "CreateSlides缺少slides数组参数"
            Return False
        End If
        
        Dim slidesArray = CType(slides, JArray)
        If slidesArray.Count = 0 Then
            errorMessage = "CreateSlides的slides数组不能为空"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateAddAnimation(params As JToken, ByRef errorMessage As String) As Boolean
        Dim effect = params("effect")?.ToString()
        If String.IsNullOrEmpty(effect) Then
            errorMessage = "AddAnimation缺少effect参数"
            Return False
        End If
        
        Dim validEffects = {"fadein", "flyin", "zoom", "wipe", "appear", "float"}
        If Not validEffects.Contains(effect.ToLower()) Then
            errorMessage = $"AddAnimation的effect参数无效: {effect}。有效值: fadeIn, flyIn, zoom, wipe, appear, float"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateApplyTransition(params As JToken, ByRef errorMessage As String) As Boolean
        Dim transType = params("transitionType")?.ToString()
        If String.IsNullOrEmpty(transType) Then
            errorMessage = "ApplyTransition缺少transitionType参数"
            Return False
        End If
        
        Dim validTypes = {"fade", "push", "wipe", "split", "reveal", "random"}
        If Not validTypes.Contains(transType.ToLower()) Then
            errorMessage = $"ApplyTransition的transitionType参数无效: {transType}。有效值: fade, push, wipe, split, reveal, random"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateBeautifySlides(params As JToken, ByRef errorMessage As String) As Boolean
        ' BeautifySlides参数都是可选的，但至少需要一个
        If params("scope") Is Nothing AndAlso params("theme") Is Nothing Then
            errorMessage = "BeautifySlides至少需要scope或theme参数"
            Return False
        End If
        Return True
    End Function

#Region "新增命令验证方法"

    Private Shared Function ValidateDeleteSlide(params As JToken, ByRef errorMessage As String) As Boolean
        If params("slideIndex") Is Nothing Then
            errorMessage = "DeleteSlide缺少slideIndex参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateDuplicateSlide(params As JToken, ByRef errorMessage As String) As Boolean
        If params("slideIndex") Is Nothing Then
            errorMessage = "DuplicateSlide缺少slideIndex参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateMoveSlide(params As JToken, ByRef errorMessage As String) As Boolean
        If params("fromIndex") Is Nothing Then
            errorMessage = "MoveSlide缺少fromIndex参数"
            Return False
        End If
        If params("toIndex") Is Nothing Then
            errorMessage = "MoveSlide缺少toIndex参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateFormatText(params As JToken, ByRef errorMessage As String) As Boolean
        ' 至少需要一个格式化属性
        If params("bold") Is Nothing AndAlso params("italic") Is Nothing AndAlso
           params("fontSize") Is Nothing AndAlso params("fontName") Is Nothing AndAlso
           params("color") Is Nothing Then
            errorMessage = "FormatText至少需要一个格式属性"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateInsertImage(params As JToken, ByRef errorMessage As String) As Boolean
        Dim imagePath = params("imagePath")?.ToString()
        If String.IsNullOrEmpty(imagePath) Then
            errorMessage = "InsertImage缺少imagePath参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateSetSlideLayout(params As JToken, ByRef errorMessage As String) As Boolean
        Dim layout = params("layout")?.ToString()
        If String.IsNullOrEmpty(layout) Then
            errorMessage = "SetSlideLayout缺少layout参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateInsertChart(params As JToken, ByRef errorMessage As String) As Boolean
        Dim chartType = params("chartType")?.ToString()
        If String.IsNullOrEmpty(chartType) Then
            errorMessage = "InsertChart缺少chartType参数"
            Return False
        End If
        If params("data") Is Nothing Then
            errorMessage = "InsertChart缺少data参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateInsertVideo(params As JToken, ByRef errorMessage As String) As Boolean
        Dim videoPath = params("videoPath")?.ToString()
        If String.IsNullOrEmpty(videoPath) Then
            errorMessage = "InsertVideo缺少videoPath参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateAddSpeakerNotes(params As JToken, ByRef errorMessage As String) As Boolean
        Dim notes = params("notes")?.ToString()
        If String.IsNullOrEmpty(notes) Then
            errorMessage = "AddSpeakerNotes缺少notes参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateSetSlideShow(params As JToken, ByRef errorMessage As String) As Boolean
        ' 所有参数都是可选的
        Return True
    End Function

    Private Shared Function ValidateApplyTheme(params As JToken, ByRef errorMessage As String) As Boolean
        ' 至少需要themeName或themeFile之一
        If params("themeName") Is Nothing AndAlso params("themeFile") Is Nothing Then
            errorMessage = "ApplyTheme需要themeName或themeFile参数"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateEditSlideMaster(params As JToken, ByRef errorMessage As String) As Boolean
        ' 至少需要一个属性
        If params("background") Is Nothing AndAlso params("titleFont") Is Nothing AndAlso params("bodyFont") Is Nothing Then
            errorMessage = "EditSlideMaster至少需要一个属性"
            Return False
        End If
        Return True
    End Function

    Private Shared Function ValidateExecuteVBA(params As JToken, ByRef errorMessage As String) As Boolean
        Dim code = params("code")?.ToString()
        If String.IsNullOrEmpty(code) Then
            errorMessage = "ExecuteVBA缺少code参数"
            Return False
        End If
        
        ' 基本的VBA代码验证
        If Not code.ToLower().Contains("sub") AndAlso Not code.ToLower().Contains("function") Then
            errorMessage = "ExecuteVBA的code必须包含Sub或Function定义"
            Return False
        End If
        
        Return True
    End Function

#End Region

End Class
