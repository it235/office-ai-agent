' ShareRibbon\Services\Reformat\TaggingValidator.vb
' AI标注结果校验器

Imports Newtonsoft.Json.Linq

''' <summary>
''' 标注校验器 - 校验AI返回的语义标注结果
''' 支持自动修复轻微错误，严重错误触发重试
''' </summary>
Public Class TaggingValidator

    ''' <summary>校验结果</summary>
    Public Class ValidationResult
        ''' <summary>是否校验通过</summary>
        Public Property IsValid As Boolean = True
        ''' <summary>错误列表</summary>
        Public Property Errors As New List(Of String)()
        ''' <summary>自动修复的项目数</summary>
        Public Property AutoFixedCount As Integer = 0
        ''' <summary>校验并修复后的标注列表</summary>
        Public Property ValidatedTags As New List(Of TaggedParagraph)()
    End Class

    ''' <summary>
    ''' 校验AI标注结果
    ''' </summary>
    ''' <param name="taggingJson">AI返回的JSON字符串</param>
    ''' <param name="mapping">语义样式映射</param>
    ''' <param name="totalParagraphs">文档总段落数</param>
    Public Shared Function Validate(
        taggingJson As String,
        mapping As SemanticStyleMapping,
        totalParagraphs As Integer) As ValidationResult

        Dim result As New ValidationResult()

        ' 解析JSON
        Dim tagsArray As JArray = Nothing
        Try
            tagsArray = TryParseTaggingJson(taggingJson)
        Catch ex As Exception
            result.IsValid = False
            result.Errors.Add($"JSON解析失败: {ex.Message}")
            Return result
        End Try

        If tagsArray Is Nothing OrElse tagsArray.Count = 0 Then
            result.IsValid = False
            result.Errors.Add("标注结果为空或格式不正确")
            Return result
        End If

        Dim availableTags = mapping.GetAvailableTagIds()
        Dim severeErrorCount As Integer = 0

        For Each item In tagsArray
            Dim paraIndex As Integer = -1
            Dim tagId As String = ""

            Try
                paraIndex = CInt(item("paraIndex"))
                tagId = If(item("tag")?.ToString(), "")
            Catch
                result.Errors.Add($"标注项格式错误: {item.ToString()}")
                severeErrorCount += 1
                Continue For
            End Try

            ' 校验paraIndex范围
            If paraIndex < 0 OrElse paraIndex >= totalParagraphs Then
                result.Errors.Add($"paraIndex越界: {paraIndex} (总共{totalParagraphs}段)")
                severeErrorCount += 1
                Continue For
            End If

            ' 校验标签是否空
            If String.IsNullOrEmpty(tagId) Then
                result.Errors.Add($"段落{paraIndex}标签为空")
                severeErrorCount += 1
                Continue For
            End If

            ' 校验标签是否合法
            If Not availableTags.Contains(tagId) Then
                ' 尝试自动修复：回退到父级
                Dim parentId = SemanticTagRegistry.GetParentTag(tagId)
                If Not String.IsNullOrEmpty(parentId) AndAlso availableTags.Contains(parentId) Then
                    result.Errors.Add($"段落{paraIndex}标签'{tagId}'不存在，自动修正为'{parentId}'")
                    tagId = parentId
                    result.AutoFixedCount += 1
                ElseIf SemanticTagRegistry.IsValidTag(tagId) Then
                    ' 标签格式合法但映射中没有定义，使用最接近的
                    Dim fallback = FindClosestTag(tagId, availableTags)
                    If Not String.IsNullOrEmpty(fallback) Then
                        result.Errors.Add($"段落{paraIndex}标签'{tagId}'不在映射中，自动修正为'{fallback}'")
                        tagId = fallback
                        result.AutoFixedCount += 1
                    Else
                        result.Errors.Add($"段落{paraIndex}使用了未知标签'{tagId}'")
                        severeErrorCount += 1
                        Continue For
                    End If
                Else
                    result.Errors.Add($"段落{paraIndex}使用了非法标签'{tagId}'")
                    severeErrorCount += 1
                    Continue For
                End If
            End If

            result.ValidatedTags.Add(New TaggedParagraph(paraIndex, tagId))
        Next

        ' 严重错误超过20%则判定为校验失败
        If severeErrorCount > totalParagraphs * 0.2 Then
            result.IsValid = False
        End If

        Return result
    End Function

    ''' <summary>尝试解析标注JSON（兼容多种格式）</summary>
    Private Shared Function TryParseTaggingJson(json As String) As JArray
        If String.IsNullOrWhiteSpace(json) Then Return Nothing

        json = json.Trim()

        ' 移除可能的markdown代码块包裹（支持 ```json, ```javascript 等）
        json = StripMarkdownCodeBlock(json)

        ' 直接是数组
        If json.StartsWith("[") Then
            Return JArray.Parse(json)
        End If

        ' 可能是包裹在对象中
        If json.StartsWith("{") Then
            Dim obj = JObject.Parse(json)
            ' 查找数组字段
            For Each prop In obj.Properties()
                If TypeOf prop.Value Is JArray Then
                    Return CType(prop.Value, JArray)
                End If
            Next
        End If

        ' 最后尝试：提取第一个 [ 到最后一个 ] 之间的内容
        Dim firstBracket = json.IndexOf("[")
        Dim lastBracket = json.LastIndexOf("]")
        If firstBracket >= 0 AndAlso lastBracket > firstBracket Then
            Dim extracted = json.Substring(firstBracket, lastBracket - firstBracket + 1)
            Try
                Return JArray.Parse(extracted)
            Catch
                ' 提取失败，返回Nothing
            End Try
        End If

        Return Nothing
    End Function

    ''' <summary>剥离markdown代码块标记</summary>
    Private Shared Function StripMarkdownCodeBlock(json As String) As String
        If String.IsNullOrWhiteSpace(json) Then Return json

        ' 处理开头的 ``` 或 ```json 等
        If json.StartsWith("```") Then
            ' 查找第一个换行符（兼容 vbLf, vbCr, vbCrLf）
            Dim firstNewline = -1
            For i = 3 To Math.Min(json.Length - 1, 50) ' 最多检查50个字符
                If json(i) = vbLf(0) OrElse json(i) = vbCr(0) Then
                    firstNewline = i
                    Exit For
                End If
            Next

            If firstNewline > 0 Then
                json = json.Substring(firstNewline + 1)
                ' 跳过可能的连续换行
                json = json.TrimStart(vbLf(0), vbCr(0))
            Else
                ' 没有换行符，直接去掉 ```xxx 前缀（可能是 ```[）
                Dim bracketPos = json.IndexOf("[")
                If bracketPos > 0 Then
                    json = json.Substring(bracketPos)
                End If
            End If
        End If

        ' 处理结尾的 ```
        If json.EndsWith("```") Then
            json = json.Substring(0, json.Length - 3)
        ElseIf json.Contains("```") Then
            ' 可能结尾的 ``` 前有换行
            Dim lastBackticks = json.LastIndexOf("```")
            If lastBackticks > 0 Then
                ' 检查 ``` 之后是否只有空白
                Dim afterBackticks = json.Substring(lastBackticks + 3).Trim()
                If String.IsNullOrEmpty(afterBackticks) Then
                    json = json.Substring(0, lastBackticks)
                End If
            End If
        End If

        Return json.Trim()
    End Function

    ''' <summary>查找最接近的可用标签</summary>
    Private Shared Function FindClosestTag(tagId As String, availableTags As List(Of String)) As String
        ' 同父级下的第一个标签
        Dim parentId = SemanticTagRegistry.GetParentTag(tagId)
        If Not String.IsNullOrEmpty(parentId) Then
            Dim sibling = availableTags.FirstOrDefault(Function(t) SemanticTagRegistry.GetParentTag(t) = parentId)
            If sibling IsNot Nothing Then Return sibling
        End If

        ' 回退到body.normal（最安全的默认值）
        If availableTags.Contains(SemanticTagRegistry.TAG_BODY_NORMAL) Then
            Return SemanticTagRegistry.TAG_BODY_NORMAL
        End If

        Return If(availableTags.FirstOrDefault(), "")
    End Function
End Class
