' ShareRibbon\Storage\PromptTemplateRepository.vb
' prompt_template 表 CRUD 与按场景/Skills 加载

Imports System.Data.SQLite
Imports Newtonsoft.Json.Linq

''' <summary>
''' Skill 输入参数定义
''' </summary>
Public Class SkillParameter
    Public Property Name As String
    Public Property Description As String
    Public Property Required As Boolean = False
    Public Property DataType As String = "string"
    Public Property DefaultValue As String
End Class

''' <summary>
''' 提示词/ Skill 记录
''' </summary>
Public Class PromptTemplateRecord
    Public Property Id As Long
    Public Property TemplateName As String
    Public Property Scenario As String
    Public Property Content As String
    Public Property IsSkill As Integer

    ''' <summary>
    ''' Skill描述（Claude规范）
    ''' </summary>
    Public Property Description As String

    ''' <summary>
    ''' Skill关键词/触发词（用于匹配）
    ''' </summary>
    Public Property Keywords As String

    ''' <summary>
    ''' Skill类别
    ''' </summary>
    Public Property Category As String

    ''' <summary>
    ''' Skill优先级（0-100，越高越优先）
    ''' </summary>
    Public Property Priority As Integer = 50

    ''' <summary>
    ''' 是否启用
    ''' </summary>
    Public Property Enabled As Boolean = True

    ''' <summary>
    ''' Skill输入参数（JSON数组）
    ''' </summary>
    Public Property ParametersJson As String

    ''' <summary>
    ''' 额外配置（兼容原有字段）
    ''' </summary>
    Public Property ExtraJson As String

    Public Property Sort As Integer
End Class

''' <summary>
''' prompt_template 表访问
''' </summary>
Public Class PromptTemplateRepository

    ''' <summary>
    ''' 按 scenario 获取系统提示词（is_skill=0）
    ''' </summary>
    Public Shared Function GetSystemPrompt(scenario As String) As String
        OfficeAiDatabase.EnsureInitialized()
        Dim scenarioNorm = If(String.IsNullOrEmpty(scenario), "common", scenario.ToLowerInvariant())
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Dim sql = "SELECT content FROM prompt_template WHERE scenario=@s AND is_skill=0 ORDER BY sort, id LIMIT 1"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@s", scenarioNorm)
                Dim obj = cmd.ExecuteScalar()
                Return If(obj Is Nothing OrElse obj Is DBNull.Value, "", obj.ToString())
            End Using
        End Using
    End Function

    ''' <summary>
    ''' 按 scenario 与 supported_apps 获取已启用 Skills（is_skill=1）
    ''' extra_json 中 supported_apps 为 JSON 数组，如 ["Excel","Word"]
    ''' </summary>
    Public Shared Function GetSkillsForApp(scenario As String, appType As String) As List(Of PromptTemplateRecord)
        OfficeAiDatabase.EnsureInitialized()
        Dim list As New List(Of PromptTemplateRecord)()
        Dim scenarioNorm = If(String.IsNullOrEmpty(scenario), "common", scenario.ToLowerInvariant())
        Dim appNorm = If(String.IsNullOrEmpty(appType), "", appType.ToLowerInvariant())

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Dim sql = "SELECT id, template_name, scenario, content, is_skill, extra_json, sort, description, keywords, category, priority, enabled, parameters_json FROM prompt_template WHERE scenario=@s AND is_skill=1 ORDER BY sort, id"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@s", scenarioNorm)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        Dim extra = If(rdr.IsDBNull(5), "", rdr.GetString(5))
                        Dim supportedApps As New List(Of String)()
                        If Not String.IsNullOrWhiteSpace(extra) Then
                            Try
                                Dim jo = JObject.Parse(extra)
                                Dim arr = If(jo("supported_apps"), jo("supportedApps"))
                                If arr IsNot Nothing AndAlso TypeOf arr Is JArray Then
                                    For Each t In CType(arr, JArray)
                                        supportedApps.Add(t.ToString().ToLowerInvariant())
                                    Next
                                End If
                            Catch
                            End Try
                        End If
                        If supportedApps.Count = 0 OrElse supportedApps.Contains(appNorm) Then
                            Dim record = New PromptTemplateRecord With {
                                .Id = rdr.GetInt64(0),
                                .TemplateName = If(rdr.IsDBNull(1), "", rdr.GetString(1)),
                                .Scenario = If(rdr.IsDBNull(2), "", rdr.GetString(2)),
                                .Content = If(rdr.IsDBNull(3), "", rdr.GetString(3)),
                                .IsSkill = rdr.GetInt32(4),
                                .ExtraJson = extra,
                                .Sort = If(rdr.IsDBNull(6), 0, rdr.GetInt32(6))
                            }
                            ' 读取新字段（兼容旧数据库）
                            If rdr.FieldCount > 7 Then
                                record.Description = If(rdr.IsDBNull(7), "", rdr.GetString(7))
                                record.Keywords = If(rdr.IsDBNull(8), "", rdr.GetString(8))
                                record.Category = If(rdr.IsDBNull(9), "", rdr.GetString(9))
                                record.Priority = If(rdr.IsDBNull(10), 50, rdr.GetInt32(10))
                                record.Enabled = If(rdr.IsDBNull(11), 1, rdr.GetInt32(11)) = 1
                                record.ParametersJson = If(rdr.IsDBNull(12), "", rdr.GetString(12))
                            End If
                            list.Add(record)
                        End If
                    End While
                End Using
            End Using
        End Using

        Return list
    End Function

    ''' <summary>
    ''' 按 scenario 列出所有记录（系统提示词 + Skills）
    ''' </summary>
    Public Shared Function ListByScenario(scenario As String) As List(Of PromptTemplateRecord)
        OfficeAiDatabase.EnsureInitialized()
        Dim list As New List(Of PromptTemplateRecord)()
        Dim scenarioNorm = If(String.IsNullOrEmpty(scenario), "common", scenario.ToLowerInvariant())
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Dim sql = "SELECT id, template_name, scenario, content, is_skill, extra_json, sort, description, keywords, category, priority, enabled, parameters_json FROM prompt_template WHERE scenario=@s ORDER BY is_skill, sort, id"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@s", scenarioNorm)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        Dim record = New PromptTemplateRecord With {
                            .Id = rdr.GetInt64(0),
                            .TemplateName = If(rdr.IsDBNull(1), "", rdr.GetString(1)),
                            .Scenario = If(rdr.IsDBNull(2), "", rdr.GetString(2)),
                            .Content = If(rdr.IsDBNull(3), "", rdr.GetString(3)),
                            .IsSkill = rdr.GetInt32(4),
                            .ExtraJson = If(rdr.IsDBNull(5), "", rdr.GetString(5)),
                            .Sort = If(rdr.IsDBNull(6), 0, rdr.GetInt32(6))
                        }
                        ' 读取新字段（兼容旧数据库）
                        If rdr.FieldCount > 7 Then
                            record.Description = If(rdr.IsDBNull(7), "", rdr.GetString(7))
                            record.Keywords = If(rdr.IsDBNull(8), "", rdr.GetString(8))
                            record.Category = If(rdr.IsDBNull(9), "", rdr.GetString(9))
                            record.Priority = If(rdr.IsDBNull(10), 50, rdr.GetInt32(10))
                            record.Enabled = If(rdr.IsDBNull(11), 1, rdr.GetInt32(11)) = 1
                            record.ParametersJson = If(rdr.IsDBNull(12), "", rdr.GetString(12))
                        End If
                        list.Add(record)
                    End While
                End Using
            End Using
        End Using
        Return list
    End Function

    ''' <summary>
    ''' 插入新记录
    ''' </summary>
    Public Shared Function Insert(record As PromptTemplateRecord) As Long
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Dim sql = "INSERT INTO prompt_template (template_name, scenario, content, is_skill, extra_json, sort, description, keywords, category, priority, enabled, parameters_json) VALUES (@name, @scenario, @content, @iskill, @extra, @sort, @desc, @keywords, @category, @priority, @enabled, @params); SELECT last_insert_rowid();"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@name", If(record.TemplateName, ""))
                cmd.Parameters.AddWithValue("@scenario", If(record.Scenario, "common"))
                cmd.Parameters.AddWithValue("@content", If(record.Content, ""))
                cmd.Parameters.AddWithValue("@iskill", record.IsSkill)
                cmd.Parameters.AddWithValue("@extra", If(record.ExtraJson, ""))
                cmd.Parameters.AddWithValue("@sort", record.Sort)
                cmd.Parameters.AddWithValue("@desc", If(record.Description, ""))
                cmd.Parameters.AddWithValue("@keywords", If(record.Keywords, ""))
                cmd.Parameters.AddWithValue("@category", If(record.Category, ""))
                cmd.Parameters.AddWithValue("@priority", record.Priority)
                cmd.Parameters.AddWithValue("@enabled", If(record.Enabled, 1, 0))
                cmd.Parameters.AddWithValue("@params", If(record.ParametersJson, ""))
                Return Convert.ToInt64(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    ''' <summary>
    ''' 更新记录
    ''' </summary>
    Public Shared Sub Update(record As PromptTemplateRecord)
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Dim sql = "UPDATE prompt_template SET template_name=@name, scenario=@scenario, content=@content, is_skill=@iskill, extra_json=@extra, sort=@sort, description=@desc, keywords=@keywords, category=@category, priority=@priority, enabled=@enabled, parameters_json=@params, update_time=datetime('now','localtime') WHERE id=@id"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@id", record.Id)
                cmd.Parameters.AddWithValue("@name", If(record.TemplateName, ""))
                cmd.Parameters.AddWithValue("@scenario", If(record.Scenario, "common"))
                cmd.Parameters.AddWithValue("@content", If(record.Content, ""))
                cmd.Parameters.AddWithValue("@iskill", record.IsSkill)
                cmd.Parameters.AddWithValue("@extra", If(record.ExtraJson, ""))
                cmd.Parameters.AddWithValue("@sort", record.Sort)
                cmd.Parameters.AddWithValue("@desc", If(record.Description, ""))
                cmd.Parameters.AddWithValue("@keywords", If(record.Keywords, ""))
                cmd.Parameters.AddWithValue("@category", If(record.Category, ""))
                cmd.Parameters.AddWithValue("@priority", record.Priority)
                cmd.Parameters.AddWithValue("@enabled", If(record.Enabled, 1, 0))
                cmd.Parameters.AddWithValue("@params", If(record.ParametersJson, ""))
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 删除记录
    ''' </summary>
    Public Shared Sub Delete(id As Long)
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand("DELETE FROM prompt_template WHERE id=@id", conn)
                cmd.Parameters.AddWithValue("@id", id)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 变量替换：{{变量名}} 替换为 vars 字典中的值
    ''' </summary>
    Public Shared Function ReplaceVariables(template As String, vars As Dictionary(Of String, String)) As String
        If String.IsNullOrEmpty(template) Then Return ""
        If vars Is Nothing OrElse vars.Count = 0 Then Return template

        Dim result = template
        For Each kv In vars
            result = result.Replace("{{" & kv.Key & "}}", If(kv.Value, ""))
        Next
        Return result
    End Function
End Class
