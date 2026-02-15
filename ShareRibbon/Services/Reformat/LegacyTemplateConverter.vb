' ShareRibbon\Services\Reformat\LegacyTemplateConverter.vb
' 旧模板 → SemanticStyleMapping 转换器

''' <summary>
''' 旧模板转换器 - 将现有ReformatTemplate转换为SemanticStyleMapping
''' 提供向后兼容能力
''' </summary>
Public Class LegacyTemplateConverter

    ''' <summary>
    ''' 将ReformatTemplate转换为SemanticStyleMapping
    ''' </summary>
    Public Shared Function Convert(template As ReformatTemplate) As SemanticStyleMapping
        If template Is Nothing Then Return Nothing

        ' 先检查缓存
        Dim cached = SemanticMappingManager.Instance.GetMappingBySourceId(template.Id)
        If cached IsNot Nothing Then Return cached

        Dim mapping As New SemanticStyleMapping()
        mapping.Name = template.Name
        mapping.SourceType = SemanticMappingSourceType.FromLegacy
        mapping.SourceId = template.Id

        ' 转换正文样式规则
        ConvertBodyStyles(template.BodyStyles, mapping)

        ' 转换版式骨架
        ConvertLayout(template.Layout, mapping)

        ' 转换页面设置（直接复用）
        If template.PageSettings IsNot Nothing Then
            mapping.PageConfig = template.PageSettings
        End If

        ' 确保基础标签
        EnsureBasicTags(mapping)

        ' 缓存转换结果
        SemanticMappingManager.Instance.AddMapping(mapping)

        Return mapping
    End Function

    ''' <summary>转换正文样式规则为语义标签</summary>
    Private Shared Sub ConvertBodyStyles(styles As List(Of StyleRule), mapping As SemanticStyleMapping)
        If styles Is Nothing Then Return

        For Each rule In styles
            Dim tag = MapRuleToTag(rule)
            If tag IsNot Nothing Then
                mapping.SemanticTags.Add(tag)
            End If
        Next
    End Sub

    ''' <summary>将StyleRule映射到SemanticTag</summary>
    Private Shared Function MapRuleToTag(rule As StyleRule) As SemanticTag
        If rule Is Nothing Then Return Nothing

        Dim tagId As String = ""
        Dim displayName As String = rule.RuleName
        Dim parentId As String = ""
        Dim matchHint As String = rule.MatchCondition

        Dim name = If(rule.RuleName, "").ToLower()

        Select Case True
            Case name.Contains("一级") OrElse name.Contains("大标题") OrElse name.Contains("章标题")
                tagId = SemanticTagRegistry.TAG_TITLE_1
                parentId = SemanticTagRegistry.TAG_TITLE

            Case name.Contains("二级") OrElse name.Contains("节标题")
                tagId = SemanticTagRegistry.TAG_TITLE_2
                parentId = SemanticTagRegistry.TAG_TITLE

            Case name.Contains("三级") OrElse name.Contains("小标题")
                tagId = SemanticTagRegistry.TAG_TITLE_3
                parentId = SemanticTagRegistry.TAG_TITLE

            Case name.Contains("正文") OrElse name.Contains("body")
                tagId = SemanticTagRegistry.TAG_BODY_NORMAL
                parentId = SemanticTagRegistry.TAG_BODY

            Case name.Contains("强调") OrElse name.Contains("emphasis")
                tagId = SemanticTagRegistry.TAG_BODY_EMPHASIS
                parentId = SemanticTagRegistry.TAG_BODY

            Case name.Contains("列表") OrElse name.Contains("list")
                tagId = SemanticTagRegistry.TAG_LIST_ORDERED
                parentId = SemanticTagRegistry.TAG_LIST

            Case name.Contains("引用") OrElse name.Contains("quote")
                tagId = SemanticTagRegistry.TAG_QUOTE
                parentId = ""

            Case name.Contains("题注") OrElse name.Contains("caption")
                tagId = SemanticTagRegistry.TAG_CAPTION
                parentId = ""

            Case Else
                ' 无法识别的规则作为正文处理
                tagId = SemanticTagRegistry.TAG_BODY_NORMAL
                parentId = SemanticTagRegistry.TAG_BODY
                matchHint = $"来自旧规则: {rule.RuleName}"
        End Select

        ' 避免重复
        Dim tag As New SemanticTag(tagId, displayName, parentId, SemanticTagRegistry.GetTagLevel(tagId), matchHint)

        ' 复制格式配置
        If rule.Font IsNot Nothing Then tag.Font = rule.Font
        If rule.Paragraph IsNot Nothing Then tag.Paragraph = rule.Paragraph
        If rule.Color IsNot Nothing Then tag.Color = rule.Color

        Return tag
    End Function

    ''' <summary>转换版式骨架</summary>
    Private Shared Sub ConvertLayout(layout As LayoutConfig, mapping As SemanticStyleMapping)
        If layout Is Nothing OrElse layout.Elements Is Nothing Then Return
        mapping.LayoutSkeleton = layout
    End Sub

    ''' <summary>确保基础标签存在</summary>
    Private Shared Sub EnsureBasicTags(mapping As SemanticStyleMapping)
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = SemanticTagRegistry.TAG_BODY_NORMAL) Then
            mapping.SemanticTags.Add(New SemanticTag(
                SemanticTagRegistry.TAG_BODY_NORMAL, "正文",
                SemanticTagRegistry.TAG_BODY, 2, "普通正文段落"))
        End If

        If Not mapping.SemanticTags.Any(Function(t) t.TagId = SemanticTagRegistry.TAG_TITLE_1) Then
            mapping.SemanticTags.Add(New SemanticTag(
                SemanticTagRegistry.TAG_TITLE_1, "一级标题",
                SemanticTagRegistry.TAG_TITLE, 2, "主要章节标题"))
        End If
    End Sub
End Class
