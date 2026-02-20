' ShareRibbon\Storage\FormatTemplateRepository.vb
' 排版模板数据仓库

Imports System.Data.SQLite
Imports System.IO
Imports Newtonsoft.Json

''' <summary>
''' 排版模板数据仓库
''' </summary>
Public Class FormatTemplateRepository
    Public Sub New()
        OfficeAiDatabase.EnsureInitialized()
    End Sub

    ''' <summary>
    ''' 获取所有模板
    ''' </summary>
    Public Function GetAllTemplates() As List(Of ReformatTemplate)
        Dim templates As New List(Of ReformatTemplate)()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand("SELECT * FROM format_template ORDER BY created_at DESC", conn)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        templates.Add(ReadTemplateFromReader(reader))
                    End While
                End Using
            End Using

            For Each template In templates
                LoadTemplateElements(template, conn)
                LoadTemplateStyleRules(template, conn)
            Next
        End Using
        Return templates
    End Function

    ''' <summary>
    ''' 根据应用类型获取模板
    ''' </summary>
    Public Function GetTemplatesByApp(targetApp As String) As List(Of ReformatTemplate)
        Dim templates As New List(Of ReformatTemplate)()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand("SELECT * FROM format_template WHERE target_app = @targetApp OR target_app = '全部' ORDER BY created_at DESC", conn)
                cmd.Parameters.AddWithValue("@targetApp", targetApp)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        templates.Add(ReadTemplateFromReader(reader))
                    End While
                End Using
            End Using

            For Each template In templates
                LoadTemplateElements(template, conn)
                LoadTemplateStyleRules(template, conn)
            Next
        End Using
        Return templates
    End Function

    ''' <summary>
    ''' 根据ID获取模板
    ''' </summary>
    Public Function GetTemplateById(templateId As String) As ReformatTemplate
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand("SELECT * FROM format_template WHERE template_id = @templateId", conn)
                cmd.Parameters.AddWithValue("@templateId", templateId)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Dim template = ReadTemplateFromReader(reader)
                        LoadTemplateElements(template, conn)
                        LoadTemplateStyleRules(template, conn)
                        Return template
                    End If
                End Using
            End Using
        End Using
        Return Nothing
    End Function

    ''' <summary>
    ''' 保存或更新模板
    ''' </summary>
    Public Sub SaveTemplate(template As ReformatTemplate)
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using transaction = conn.BeginTransaction()
                Try
                    Dim existing = GetTemplateById(template.Id)
                    If existing Is Nothing Then
                        InsertTemplate(template, conn)
                    Else
                        UpdateTemplate(template, conn)
                    End If

                    DeleteTemplateElements(template.Id, conn)
                    InsertTemplateElements(template, conn)

                    DeleteTemplateStyleRules(template.Id, conn)
                    InsertTemplateStyleRules(template, conn)

                    transaction.Commit()
                Catch
                    transaction.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 删除模板
    ''' </summary>
    Public Function DeleteTemplate(templateId As String) As Boolean
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand("DELETE FROM format_template WHERE template_id = @templateId", conn)
                cmd.Parameters.AddWithValue("@templateId", templateId)
                Return cmd.ExecuteNonQuery() > 0
            End Using
        End Using
    End Function

    ''' <summary>
    ''' 获取源文件Blob
    ''' </summary>
    Public Function GetSourceFileBlob(templateId As String) As Byte()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand("SELECT source_file_blob FROM format_template WHERE template_id = @templateId", conn)
                cmd.Parameters.AddWithValue("@templateId", templateId)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    Return DirectCast(result, Byte())
                End If
            End Using
        End Using
        Return Nothing
    End Function

    ''' <summary>
    ''' 保存源文件Blob
    ''' </summary>
    Public Sub SaveSourceFileBlob(templateId As String, fileBytes As Byte())
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand("UPDATE format_template SET source_file_blob = @blob WHERE template_id = @templateId", conn)
                cmd.Parameters.AddWithValue("@templateId", templateId)
                cmd.Parameters.AddWithValue("@blob", fileBytes)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Function ReadTemplateFromReader(reader As SQLiteDataReader) As ReformatTemplate
        Dim template As New ReformatTemplate With {
            .Id = If(IsDBNull(reader("template_id")), "", reader("template_id").ToString()),
            .Name = If(IsDBNull(reader("name")), "", reader("name").ToString()),
            .Description = If(IsDBNull(reader("description")), "", reader("description").ToString()),
            .Category = If(IsDBNull(reader("category")), "通用", reader("category").ToString()),
            .TargetApp = If(IsDBNull(reader("target_app")), "Word", reader("target_app").ToString()),
            .IsPreset = If(IsDBNull(reader("is_preset")), False, Convert.ToBoolean(reader("is_preset"))),
            .AiGuidance = If(IsDBNull(reader("ai_guidance")), "", reader("ai_guidance").ToString()),
            .ThumbnailBase64 = If(IsDBNull(reader("thumbnail_base64")), "", reader("thumbnail_base64").ToString()),
            .SourceFileName = If(IsDBNull(reader("source_file_name")), "", reader("source_file_name").ToString()),
            .SourceFileContent = If(IsDBNull(reader("source_file_content")), "", reader("source_file_content").ToString()),
            .CreatedAt = If(IsDBNull(reader("created_at")), DateTime.Now, DateTime.Parse(reader("created_at").ToString())),
            .LastModified = If(IsDBNull(reader("last_modified")), DateTime.Now, DateTime.Parse(reader("last_modified").ToString()))
        }

        Dim sourceStr = If(IsDBNull(reader("template_source")), "manual", reader("template_source").ToString())
        If [Enum].TryParse(Of TemplateSourceType)(sourceStr, True, template.TemplateSource) Then
        End If

        If Not IsDBNull(reader("page_settings_json")) Then
            template.PageSettings = JsonConvert.DeserializeObject(Of PageConfig)(reader("page_settings_json").ToString())
        End If

        Return template
    End Function

    Private Sub LoadTemplateElements(template As ReformatTemplate, conn As SQLiteConnection)
        Using cmd As New SQLiteCommand("SELECT * FROM format_element WHERE template_id = @templateId ORDER BY sort_order", conn)
            cmd.Parameters.AddWithValue("@templateId", template.Id)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim element As New LayoutElement With {
                        .Name = If(IsDBNull(reader("element_name")), "", reader("element_name").ToString()),
                        .ElementType = If(IsDBNull(reader("element_type")), "text", reader("element_type").ToString()),
                        .DefaultValue = If(IsDBNull(reader("default_value")), "", reader("default_value").ToString()),
                        .Required = If(IsDBNull(reader("is_required")), True, Convert.ToBoolean(reader("is_required"))),
                        .SortOrder = If(IsDBNull(reader("sort_order")), 0, Convert.ToInt32(reader("sort_order"))),
                        .PlaceholderContent = If(IsDBNull(reader("placeholder_content")), "{{content}}", reader("placeholder_content").ToString())
                    }

                    If Not IsDBNull(reader("font_config_json")) Then
                        element.Font = JsonConvert.DeserializeObject(Of FontConfig)(reader("font_config_json").ToString())
                    End If
                    If Not IsDBNull(reader("paragraph_config_json")) Then
                        element.Paragraph = JsonConvert.DeserializeObject(Of ParagraphConfig)(reader("paragraph_config_json").ToString())
                    End If
                    If Not IsDBNull(reader("color_config_json")) Then
                        element.Color = JsonConvert.DeserializeObject(Of ColorConfig)(reader("color_config_json").ToString())
                    End If
                    If Not IsDBNull(reader("special_props_json")) Then
                        element.SpecialProps = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(reader("special_props_json").ToString())
                    End If

                    template.Layout.Elements.Add(element)
                End While
            End Using
        End Using
    End Sub

    Private Sub LoadTemplateStyleRules(template As ReformatTemplate, conn As SQLiteConnection)
        Using cmd As New SQLiteCommand("SELECT * FROM format_style_rule WHERE template_id = @templateId ORDER BY sort_order", conn)
            cmd.Parameters.AddWithValue("@templateId", template.Id)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim rule As New StyleRule With {
                        .RuleName = If(IsDBNull(reader("rule_name")), "", reader("rule_name").ToString()),
                        .MatchCondition = If(IsDBNull(reader("match_condition")), "", reader("match_condition").ToString()),
                        .SortOrder = If(IsDBNull(reader("sort_order")), 0, Convert.ToInt32(reader("sort_order")))
                    }

                    If Not IsDBNull(reader("font_config_json")) Then
                        rule.Font = JsonConvert.DeserializeObject(Of FontConfig)(reader("font_config_json").ToString())
                    End If
                    If Not IsDBNull(reader("paragraph_config_json")) Then
                        rule.Paragraph = JsonConvert.DeserializeObject(Of ParagraphConfig)(reader("paragraph_config_json").ToString())
                    End If
                    If Not IsDBNull(reader("color_config_json")) Then
                        rule.Color = JsonConvert.DeserializeObject(Of ColorConfig)(reader("color_config_json").ToString())
                    End If

                    template.BodyStyles.Add(rule)
                End While
            End Using
        End Using
    End Sub

    Private Sub InsertTemplate(template As ReformatTemplate, conn As SQLiteConnection)
        Using cmd As New SQLiteCommand(
            "INSERT INTO format_template (template_id, name, description, category, target_app, is_preset, template_source, source_file_name, source_file_content, layout_json, style_rules_json, page_settings_json, ai_guidance, thumbnail_base64, created_at, last_modified) " &
            "VALUES (@templateId, @name, @desc, @category, @targetApp, @isPreset, @source, @sourceFileName, @sourceFileContent, @layoutJson, @styleJson, @pageJson, @aiGuidance, @thumbnail, @createdAt, @lastModified)", conn)
            cmd.Parameters.AddWithValue("@templateId", template.Id)
            cmd.Parameters.AddWithValue("@name", template.Name)
            cmd.Parameters.AddWithValue("@desc", template.Description)
            cmd.Parameters.AddWithValue("@category", template.Category)
            cmd.Parameters.AddWithValue("@targetApp", template.TargetApp)
            cmd.Parameters.AddWithValue("@isPreset", template.IsPreset)
            cmd.Parameters.AddWithValue("@source", template.TemplateSource.ToString())
            cmd.Parameters.AddWithValue("@sourceFileName", template.SourceFileName)
            cmd.Parameters.AddWithValue("@sourceFileContent", template.SourceFileContent)
            cmd.Parameters.AddWithValue("@layoutJson", JsonConvert.SerializeObject(template.Layout))
            cmd.Parameters.AddWithValue("@styleJson", JsonConvert.SerializeObject(template.BodyStyles))
            cmd.Parameters.AddWithValue("@pageJson", JsonConvert.SerializeObject(template.PageSettings))
            cmd.Parameters.AddWithValue("@aiGuidance", template.AiGuidance)
            cmd.Parameters.AddWithValue("@thumbnail", template.ThumbnailBase64)
            cmd.Parameters.AddWithValue("@createdAt", template.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss"))
            cmd.Parameters.AddWithValue("@lastModified", template.LastModified.ToString("yyyy-MM-dd HH:mm:ss"))
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub UpdateTemplate(template As ReformatTemplate, conn As SQLiteConnection)
        Using cmd As New SQLiteCommand(
            "UPDATE format_template SET name = @name, description = @desc, category = @category, target_app = @targetApp, " &
            "is_preset = @isPreset, template_source = @source, source_file_name = @sourceFileName, source_file_content = @sourceFileContent, " &
            "layout_json = @layoutJson, style_rules_json = @styleJson, page_settings_json = @pageJson, ai_guidance = @aiGuidance, " &
            "thumbnail_base64 = @thumbnail, last_modified = @lastModified WHERE template_id = @templateId", conn)
            cmd.Parameters.AddWithValue("@templateId", template.Id)
            cmd.Parameters.AddWithValue("@name", template.Name)
            cmd.Parameters.AddWithValue("@desc", template.Description)
            cmd.Parameters.AddWithValue("@category", template.Category)
            cmd.Parameters.AddWithValue("@targetApp", template.TargetApp)
            cmd.Parameters.AddWithValue("@isPreset", template.IsPreset)
            cmd.Parameters.AddWithValue("@source", template.TemplateSource.ToString())
            cmd.Parameters.AddWithValue("@sourceFileName", template.SourceFileName)
            cmd.Parameters.AddWithValue("@sourceFileContent", template.SourceFileContent)
            cmd.Parameters.AddWithValue("@layoutJson", JsonConvert.SerializeObject(template.Layout))
            cmd.Parameters.AddWithValue("@styleJson", JsonConvert.SerializeObject(template.BodyStyles))
            cmd.Parameters.AddWithValue("@pageJson", JsonConvert.SerializeObject(template.PageSettings))
            cmd.Parameters.AddWithValue("@aiGuidance", template.AiGuidance)
            cmd.Parameters.AddWithValue("@thumbnail", template.ThumbnailBase64)
            cmd.Parameters.AddWithValue("@lastModified", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub DeleteTemplateElements(templateId As String, conn As SQLiteConnection)
        Using cmd As New SQLiteCommand("DELETE FROM format_element WHERE template_id = @templateId", conn)
            cmd.Parameters.AddWithValue("@templateId", templateId)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub InsertTemplateElements(template As ReformatTemplate, conn As SQLiteConnection)
        For Each element In template.Layout.Elements
            Using cmd As New SQLiteCommand(
                "INSERT INTO format_element (template_id, element_name, element_type, default_value, is_required, sort_order, font_config_json, paragraph_config_json, color_config_json, special_props_json, placeholder_content, is_enabled) " &
                "VALUES (@templateId, @name, @type, @default, @required, @sort, @fontJson, @paraJson, @colorJson, @specialJson, @placeholder, 1)", conn)
                cmd.Parameters.AddWithValue("@templateId", template.Id)
                cmd.Parameters.AddWithValue("@name", element.Name)
                cmd.Parameters.AddWithValue("@type", element.ElementType)
                cmd.Parameters.AddWithValue("@default", element.DefaultValue)
                cmd.Parameters.AddWithValue("@required", element.Required)
                cmd.Parameters.AddWithValue("@sort", element.SortOrder)
                cmd.Parameters.AddWithValue("@fontJson", JsonConvert.SerializeObject(element.Font))
                cmd.Parameters.AddWithValue("@paraJson", JsonConvert.SerializeObject(element.Paragraph))
                cmd.Parameters.AddWithValue("@colorJson", JsonConvert.SerializeObject(element.Color))
                cmd.Parameters.AddWithValue("@specialJson", JsonConvert.SerializeObject(element.SpecialProps))
                cmd.Parameters.AddWithValue("@placeholder", element.PlaceholderContent)
                cmd.ExecuteNonQuery()
            End Using
        Next
    End Sub

    Private Sub DeleteTemplateStyleRules(templateId As String, conn As SQLiteConnection)
        Using cmd As New SQLiteCommand("DELETE FROM format_style_rule WHERE template_id = @templateId", conn)
            cmd.Parameters.AddWithValue("@templateId", templateId)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub InsertTemplateStyleRules(template As ReformatTemplate, conn As SQLiteConnection)
        For Each rule In template.BodyStyles
            Using cmd As New SQLiteCommand(
                "INSERT INTO format_style_rule (template_id, rule_name, match_condition, sort_order, font_config_json, paragraph_config_json, color_config_json, is_enabled) " &
                "VALUES (@templateId, @name, @condition, @sort, @fontJson, @paraJson, @colorJson, 1)", conn)
                cmd.Parameters.AddWithValue("@templateId", template.Id)
                cmd.Parameters.AddWithValue("@name", rule.RuleName)
                cmd.Parameters.AddWithValue("@condition", rule.MatchCondition)
                cmd.Parameters.AddWithValue("@sort", rule.SortOrder)
                cmd.Parameters.AddWithValue("@fontJson", JsonConvert.SerializeObject(rule.Font))
                cmd.Parameters.AddWithValue("@paraJson", JsonConvert.SerializeObject(rule.Paragraph))
                cmd.Parameters.AddWithValue("@colorJson", JsonConvert.SerializeObject(rule.Color))
                cmd.ExecuteNonQuery()
            End Using
        Next
    End Sub
End Class
